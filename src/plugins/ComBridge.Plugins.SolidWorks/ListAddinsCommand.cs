using ComBridge.Core;
using Microsoft.Win32;
using SolidWorks.Interop.sldworks;

namespace ComBridge.Plugins.SolidWorks;

/// <summary>
/// <c>solidworks list-addins</c> — enumerate every registered SolidWorks
/// add-in (UI-visible + hidden product-feature) as TSV, with each
/// add-in's enabled-at-startup state and currently-loaded state in the
/// attached session.
/// </summary>
/// <remarks>
/// <para>
/// SolidWorks has no first-class "<c>GetAddIns()</c>" enumeration API.
/// The canonical answer is a registry walk plus per-add-in probes. The
/// authoritative source for this code's structure is
/// <c>C:\personal_rag\solidworks\lesson_20260529_sw_addin_dual_registry.md</c>
/// (corroborated against SOLIDWORKS 2026 SP1.1 build 34.111.0011).
/// </para>
/// <para>
/// Three registry locations matter:
/// </para>
/// <list type="bullet">
///   <item><c>HKLM\SOFTWARE\SolidWorks\AddIns\{clsid}</c> — UI-visible
///         add-ins (the ones that appear in Tools &gt; Add-Ins). Subkey
///         name is the CLSID; <c>(Default)</c> REG_SZ value is the
///         friendly name. Typically ~9-10 entries on a stock install.</item>
///   <item><c>HKLM\SOFTWARE\SolidWorks\SOLIDWORKS &lt;ver&gt;\Addins\{clsid}</c>
///         — version-specific HIDDEN add-ins (Design Checker, Costing,
///         TolAnalyst, Sustainability, ScanTo3D/Reveng, etc.). The
///         Tools &gt; Add-Ins UI does NOT show these. Same field shape
///         as the UI-visible tree. We walk every <c>SOLIDWORKS &lt;ver&gt;</c>
///         subkey we find, so multi-version SW installs surface every
///         version's hidden tree.</item>
///   <item><c>HKCU\Software\SolidWorks\AddInsStartup\{clsid}</c> — per-user
///         enabled-at-startup state. <c>(Default)</c> is REG_DWORD:
///         <c>1</c> = auto-load at next SW start, <c>0</c> = don't.
///         A missing key is effectively zero. <b>Note:</b> the lesson
///         calls out this is NOT under the AddIns path — it lives at
///         <c>AddInsStartup</c>, a sibling key with the SAME GUID format.
///         An earlier draft of this code had the wrong path.</item>
/// </list>
/// <para>
/// DLL-path resolution: <c>HKCR\CLSID\{clsid}\InprocServer32\(Default)</c>
/// gives the host binary. For .NET-hosted add-ins (which is most of
/// them in modern SW), that value is literally <c>mscoree.dll</c> — the
/// real assembly path is on the same key as a <c>CodeBase</c> named
/// value, formatted as <c>file:///C:/...</c>. We detect the
/// mscoree indirection and resolve through CodeBase automatically.
/// </para>
/// <para>
/// Currently-loaded probe: <c>ISldWorks.GetAddInObject(Clsid)</c>. The
/// canonical RAG (<c>sldworks_methods_v3_llm.rag</c> line 82) confirms
/// the parameter is the CLSID, NOT the ProgID — an earlier draft of
/// this code passed the wrong identifier. Returns the live IDispatch
/// when loaded, null otherwise. Some 3rd-party add-ins throw on this
/// probe; we tolerate with a per-add-in try/catch.
/// </para>
/// <para>
/// What this does NOT enumerate: SW also hardcodes a small set of
/// "foundational" modules into <c>SLDWORKS.EXE</c> itself
/// (<c>fworks.dll</c>/FeatureWorks, <c>swbrowser.dll</c>/Toolbox Browser,
/// <c>SwLoaderSw.dll</c>, <c>BendSequenceSwu.dll</c>). They appear in
/// neither registry tree and can't be toggled at all. They're not
/// add-ins in the user-facing sense, so omitting them is correct.
/// </para>
/// <para>
/// Output: one TSV row per distinct CLSID (case-insensitive merge).
/// Columns: <c>name | clsid | dll_path | enabled_startup | currently_loaded | scope | type</c>.
/// The <c>scope</c> column may be <c>ui-visible</c>, <c>hidden:2026</c>,
/// or <c>ui-visible;hidden:2026</c> if the CLSID appears in multiple
/// trees (rare but legal).
/// </para>
/// </remarks>
internal sealed class SwListAddinsCommand : IBridgeCommand
{
    public string Name => "list-addins";
    public string Usage => "list-addins   (lists UI-visible + hidden SW add-ins with enabled/loaded state as TSV)";

    private const string UiVisibleAddinsPath = @"SOFTWARE\SolidWorks\AddIns";
    private const string StartupPath         = @"Software\SolidWorks\AddInsStartup";
    private const string SwRoot              = @"SOFTWARE\SolidWorks";

    public Task<int> RunAsync(object comRoot, string[] args, TextWriter output)
    {
        var app = (ISldWorks)comRoot;
        output.WriteLine("# columns: name\tclsid\tdll_path\tenabled_startup\tcurrently_loaded\tscope\ttype");

        // Case-insensitive CLSID dedupe — registry GUID casing varies across
        // hives (lesson_20260529 explicitly calls this out).
        var rows = new Dictionary<string, AddinRow>(StringComparer.OrdinalIgnoreCase);

        // 1. UI-visible tree (HKLM\...\AddIns\{guid}).
        ReadAddinsTree(Registry.LocalMachine, UiVisibleAddinsPath, scope: "ui-visible", rows);

        // 2. Hidden version-specific trees
        //    (HKLM\...\SOLIDWORKS <ver>\Addins\{guid}).
        ReadHiddenTrees(rows);

        // 3. Per-user enabled-at-startup (HKCU\...\AddInsStartup\{guid}).
        ReadStartupFlags(rows);

        int count = 0;
        foreach (var row in rows.Values)
        {
            // Live-load probe via CLSID — NOT ProgID (sldworks_methods_v3_llm.rag:82).
            bool loaded = false;
            try
            {
                var obj = app.GetAddInObject(row.Clsid);
                loaded = obj is not null;
            }
            catch { loaded = false; }

            output.WriteLine(string.Join("\t",
                EscTab(row.Name),
                row.Clsid,
                EscTab(row.DllPath),
                row.EnabledAtStartup.ToString().ToLowerInvariant(),
                loaded.ToString().ToLowerInvariant(),
                row.Scope,
                row.AddinType));
            count++;
        }

        output.WriteLine();
        output.WriteLine($"# total: {count} (UI-visible + hidden; SLDWORKS.EXE-hardcoded modules are not enumerable)");
        return Task.FromResult(0);
    }

    /// <summary>
    /// Walk a single AddIns tree at the given registry path, populating
    /// or merging rows by CLSID. Adds the friendly name (HKLM Default
    /// REG_SZ), the Type field if present, and resolves the DLL path
    /// (with mscoree → CodeBase resolution for .NET hosts) on first sight.
    /// </summary>
    private static void ReadAddinsTree(RegistryKey root, string path, string scope, Dictionary<string, AddinRow> rows)
    {
        try
        {
            using var addinsKey = root.OpenSubKey(path, writable: false);
            if (addinsKey is null) return;
            foreach (var sub in addinsKey.GetSubKeyNames())
            {
                if (!LooksLikeClsid(sub)) continue;
                using var entry = addinsKey.OpenSubKey(sub, writable: false);
                if (entry is null) continue;

                if (!rows.TryGetValue(sub, out var row))
                {
                    row = new AddinRow { Clsid = sub, Scope = scope };
                    rows[sub] = row;
                }
                else if (!row.Scope.Contains(scope, StringComparison.Ordinal))
                {
                    row.Scope = row.Scope + ";" + scope;
                }

                // Friendly name comes from HKCR\CLSID\{guid}\(Default) — the
                // COM-registered class name. HKLM's AddIns subkey (Default)
                // is something else (typically the load-mode flag) and is
                // NOT the friendly name. lesson_20260529 example uses
                // `name = get_default(HKCR, fr'CLSID\{guid}')`.
                if (string.IsNullOrEmpty(row.Name))
                {
                    row.Name = ResolveFriendlyName(sub);
                }

                // Type field — 0 generic, 1 utility. Not always present;
                // some addins ship without it.
                if (entry.GetValue("Type") is int t && string.IsNullOrEmpty(row.AddinType))
                {
                    row.AddinType = t switch
                    {
                        0 => "generic",
                        1 => "utility",
                        _ => $"type{t}"
                    };
                }

                if (string.IsNullOrEmpty(row.DllPath))
                {
                    row.DllPath = ResolveDllPath(sub);
                }
            }
        }
        catch
        {
            // Tolerate missing/locked keys — UI-visible may be absent in
            // edge installs and that's a legitimate empty result.
        }
    }

    /// <summary>
    /// Find every <c>HKLM\SOFTWARE\SolidWorks\SOLIDWORKS &lt;ver&gt;\Addins</c>
    /// subkey on the machine and walk each. Multi-version SW installs (Ken
    /// has 2024 + 2026 on the same machine, for example) get every hidden
    /// tree surfaced with version-tagged scope.
    /// </summary>
    private static void ReadHiddenTrees(Dictionary<string, AddinRow> rows)
    {
        try
        {
            using var swRoot = Registry.LocalMachine.OpenSubKey(SwRoot, writable: false);
            if (swRoot is null) return;
            foreach (var sub in swRoot.GetSubKeyNames())
            {
                // Match "SOLIDWORKS 2024", "SOLIDWORKS 2026", etc. Allow only
                // the canonical "SOLIDWORKS NNNN" shape to avoid pulling in
                // sibling keys like "Setup" or "Add-Ins" itself.
                if (!sub.StartsWith("SOLIDWORKS ", StringComparison.OrdinalIgnoreCase)) continue;
                var versionToken = sub.Substring("SOLIDWORKS ".Length).Trim();
                if (versionToken.Length < 4 || !int.TryParse(versionToken[..4], out _)) continue;

                ReadAddinsTree(
                    Registry.LocalMachine,
                    $@"{SwRoot}\{sub}\Addins",
                    scope: $"hidden:{versionToken[..4]}",
                    rows);
            }
        }
        catch
        {
            // Tolerate — falling through leaves the hidden trees unsurfaced
            // but doesn't fail the command.
        }
    }

    /// <summary>
    /// Overlay HKCU per-user startup flags onto rows we've already seen
    /// from HKLM. A CLSID may legitimately exist in HKCU\AddInsStartup
    /// without an HKLM counterpart (rare — e.g. an old install whose HKLM
    /// entry was removed); in that case we add a stub row with only the
    /// startup flag and an empty name so it's still visible in the listing.
    /// </summary>
    private static void ReadStartupFlags(Dictionary<string, AddinRow> rows)
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(StartupPath, writable: false);
            if (key is null) return;
            foreach (var sub in key.GetSubKeyNames())
            {
                if (!LooksLikeClsid(sub)) continue;
                using var entry = key.OpenSubKey(sub, writable: false);
                if (entry is null) continue;

                // (Default) is REG_DWORD here, not REG_SZ — this is the
                // subtle path/type mismatch the lesson explicitly calls out.
                bool enabled = false;
                if (entry.GetValue(null) is int dword) enabled = dword != 0;

                if (rows.TryGetValue(sub, out var row))
                {
                    row.EnabledAtStartup = enabled;
                }
                else
                {
                    rows[sub] = new AddinRow
                    {
                        Clsid = sub,
                        EnabledAtStartup = enabled,
                        Scope = "hkcu-only",
                        Name = ResolveFriendlyName(sub),
                        DllPath = ResolveDllPath(sub),
                    };
                }
            }
        }
        catch { }
    }

    /// <summary>
    /// Resolve a CLSID to its COM-registered friendly name via
    /// <c>HKCR\CLSID\{clsid}\(Default)</c>. This is the class display name
    /// from the COM registration (e.g. "SWDesignCheck Class",
    /// "Picture2Sketch Class"), NOT a SW-specific value. lesson_20260529
    /// confirms this is where the SW Tools &gt; Add-Ins UI reads its
    /// labels too.
    /// </summary>
    private static string ResolveFriendlyName(string clsid)
    {
        try
        {
            using var key = Registry.ClassesRoot.OpenSubKey($@"CLSID\{clsid}", writable: false);
            return (key?.GetValue(null) as string) ?? "";
        }
        catch { return ""; }
    }

    /// <summary>
    /// Resolve a CLSID to its real DLL path via
    /// <c>HKCR\CLSID\{clsid}\InprocServer32</c>. For .NET-hosted add-ins
    /// the <c>(Default)</c> value is literally <c>mscoree.dll</c> and the
    /// actual managed assembly path lives in the <c>CodeBase</c> named
    /// value as a <c>file:///</c> URL — we detect this and resolve through.
    /// </summary>
    private static string ResolveDllPath(string clsid)
    {
        try
        {
            using var key = Registry.ClassesRoot.OpenSubKey(
                $@"CLSID\{clsid}\InprocServer32",
                writable: false);
            if (key is null) return "";

            string? path = key.GetValue(null) as string;
            if (string.IsNullOrEmpty(path)) return "";

            // .NET-hosted indirection (lesson_20260529 § ".NET-hosted addins
            // use mscoree.dll with CodeBase indirection").
            if (path.IndexOf("mscoree", StringComparison.OrdinalIgnoreCase) >= 0
                && key.GetValue("CodeBase") is string codeBase
                && codeBase.Length > 0)
            {
                if (codeBase.StartsWith("file:///", StringComparison.OrdinalIgnoreCase))
                    codeBase = codeBase.Substring("file:///".Length);
                return codeBase.Replace('/', '\\');
            }

            return path;
        }
        catch { return ""; }
    }

    /// <summary>
    /// Cheap CLSID-shape check: 38 chars, braces at the ends. Filters
    /// non-GUID subkey clutter and the occasional legacy non-add-in entry.
    /// </summary>
    private static bool LooksLikeClsid(string s)
        => s.Length == 38 && s[0] == '{' && s[^1] == '}';

    private static string EscTab(string? s)
        => (s ?? "").Replace('\t', ' ').Replace('\n', ' ').Replace('\r', ' ');

    private sealed class AddinRow
    {
        public string Clsid = "";
        public string Name = "";
        public string DllPath = "";
        public bool EnabledAtStartup;
        public string Scope = "";       // "ui-visible", "hidden:2026", "hkcu-only", or ";"-joined
        public string AddinType = "";   // "generic" / "utility" / "type<N>" / ""
    }
}
