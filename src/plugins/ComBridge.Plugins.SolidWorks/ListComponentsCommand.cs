using System.Text;
using ComBridge.Core;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace ComBridge.Plugins.SolidWorks;

/// <summary>
/// <c>solidworks list-components [path] [--config &lt;name&gt;]</c> —
/// walk the assembly's components and emit
/// <c>{"path","config","components":[{"name","path","config","suppressed"}...]}</c>.
/// </summary>
/// <remarks>
/// <para>
/// Behavior matrix (mirrors <c>list-configs</c>):
/// </para>
/// <list type="bullet">
///   <item><b>Path given, file already open</b>: walks the live doc.
///         If <c>--config</c> is supplied AND differs from the active
///         config, runs <c>ShowConfiguration2 + ForceRebuild3</c> per
///         the defense-in-depth pattern from
///         <c>lesson_20260512_opendoc6_config_arg_silent_bug.md</c>.
///         <b>This mutates the user's session</b> — the active config
///         changes. We accept that cost because "give me components
///         in config X" with a different config active is ambiguous,
///         and silently returning the wrong config's components (the
///         exact bug the lesson is about) is dangerous.</item>
///   <item><b>Path given, file NOT open</b>: passes <c>--config</c>
///         through to <c>OpenDoc6</c>'s config arg (defending against
///         the silent-config-arg trap), then verifies the active
///         config matches and runs <c>ShowConfiguration2 + ForceRebuild3</c>
///         if not.</item>
///   <item><b>Path omitted</b>: uses the active doc. <c>--config</c>
///         behaves as above.</item>
/// </list>
/// <para>
/// Safety discipline baked in:
/// </para>
/// <list type="bullet">
///   <item><c>OpenDoc6</c> always receives the actual config name
///         (never <c>""</c>) when <c>--config</c> is supplied —
///         <c>lesson_20260512</c>.</item>
///   <item>Active-config verification + <c>ShowConfiguration2</c> +
///         <c>ForceRebuild3</c> defense-in-depth — same lesson.</item>
///   <item><c>ForceRebuild3</c> runs BEFORE the component walk
///         (not mid-walk) so it doesn't invalidate component COM
///         pointers we're holding —
///         <c>lesson_20260424_forcerebuild3_invalidates_com_pointers.md</c>.</item>
///   <item>Each component RCW is hard-cast to <c>IComponent2</c>
///         inside a try/catch — NEVER via <c>as IComponent2</c> on
///         the object return —
///         <c>lesson_20260424_as_bodyfolder_cast_unreliable.md</c>.</item>
/// </list>
/// <para>
/// Exit codes: <c>0</c> success; <c>1</c> path doesn't exist, has
/// unrecognized extension, is not an assembly, or no path AND no
/// active doc.
/// </para>
/// </remarks>
internal sealed class ListComponentsCommand : IBridgeCommand
{
    public string Name => "list-components";
    public string Usage => "list-components [<path>] [--config <name>]   (lists components as JSON; omit path to use active doc)";

    public Task<int> RunAsync(object comRoot, string[] args, TextWriter output)
    {
        var app = (ISldWorks)comRoot;

        // Parse args. First non-flag positional is the path; --config
        // takes a value. Unknown flags WARN but don't fail (matches
        // the host's existing convention).
        string? path = null;
        string? requestedConfig = null;
        for (int i = 0; i < args.Length; i++)
        {
            string a = args[i];
            if (a == "--config" && i + 1 < args.Length)
            {
                requestedConfig = args[++i];
            }
            else if (a.StartsWith("--"))
            {
                output.WriteLine($"WARN: unknown flag '{a}' ignored.");
            }
            else if (path is null)
            {
                path = a;
            }
        }

        IModelDoc2? doc = null;
        bool weOpenedIt = false;
        string resolvedPath;

        if (string.IsNullOrEmpty(path))
        {
            try { doc = (IModelDoc2)app.ActiveDoc; } catch { doc = null; }
            if (doc is null)
            {
                output.WriteLine("ERROR: no path given and no active document.");
                return Task.FromResult(1);
            }
            try { resolvedPath = doc.GetPathName() ?? ""; }
            catch { resolvedPath = ""; }
        }
        else
        {
            if (!File.Exists(path))
            {
                output.WriteLine($"ERROR: file not found: {path}");
                return Task.FromResult(1);
            }
            if (SwDocSession.DocTypeFromPath(path) == 0)
            {
                output.WriteLine($"ERROR: unrecognized SOLIDWORKS file extension: {Path.GetExtension(path)}");
                return Task.FromResult(1);
            }

            // Pass the requested config (or "") through to OpenDoc6.
            // If --config wasn't supplied we use "" and accept whatever
            // SW loads; only --config triggers the silent-config-arg
            // defense.
            doc = SwDocSession.OpenForReadOrFindOpen(
                app, path, config: requestedConfig ?? "",
                out weOpenedIt, out int errs, out int warns);
            if (doc is null)
            {
                output.WriteLine($"ERROR: OpenDoc6 returned null (errors=0x{errs:X}, warnings=0x{warns:X}).");
                return Task.FromResult(1);
            }
            resolvedPath = path;
        }

        try
        {
            // Verify the doc is an assembly. The whole "walk components"
            // concept only applies to swDocASSEMBLY; parts and drawings
            // return useful but different shapes that would require a
            // separate command.
            int docType = 0;
            try { docType = doc.GetType(); } catch { }
            if (docType != (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                output.WriteLine($"ERROR: document is not an assembly (type {docType}); list-components requires .sldasm.");
                return Task.FromResult(1);
            }

            // Defense-in-depth config switch. If --config was supplied
            // AND the active config doesn't match, ShowConfiguration2 +
            // ForceRebuild3. ForceRebuild3 invalidates Feature COM
            // pointers, but we haven't fetched any yet — the component
            // walk happens AFTER this block.
            string activeConfig = "";
            try
            {
                var cfgMgr = doc.ConfigurationManager;
                if (cfgMgr is not null)
                {
                    var active = cfgMgr.ActiveConfiguration as IConfiguration;
                    if (active is not null) activeConfig = active.Name ?? "";
                }
            }
            catch { }

            if (!string.IsNullOrEmpty(requestedConfig) &&
                !string.Equals(activeConfig, requestedConfig, StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    bool shown = doc.ShowConfiguration2(requestedConfig);
                    if (!shown)
                    {
                        output.WriteLine($"WARN: ShowConfiguration2('{requestedConfig}') returned false; staying in '{activeConfig}'.");
                    }
                    else
                    {
                        // TopOnly=false so subassembly per-config
                        // suppression states re-resolve.
                        try { doc.ForceRebuild3(false); }
                        catch (Exception ex)
                        {
                            output.WriteLine($"WARN: ForceRebuild3 failed: {ex.Message}");
                        }
                        activeConfig = requestedConfig;
                    }
                }
                catch (Exception ex)
                {
                    output.WriteLine($"WARN: config switch to '{requestedConfig}' failed: {ex.Message}");
                }
            }

            // Hard-cast the assembly interface — IModelDoc2 → IAssemblyDoc.
            // The doc is already known to be swDocASSEMBLY (we checked
            // GetType above), so the cast should succeed; we still
            // try/catch for robustness.
            IAssemblyDoc? asm = null;
            try { asm = (IAssemblyDoc)doc; } catch { asm = null; }
            if (asm is null)
            {
                output.WriteLine("ERROR: failed to cast IModelDoc2 to IAssemblyDoc.");
                return Task.FromResult(1);
            }

            // GetComponents(false) = ALL components (recursive), not
            // just top-level. Returns an object that materializes as
            // object[] in C# interop; each element is a Component2 RCW.
            object[] rawComponents = Array.Empty<object>();
            try
            {
                var raw = asm.GetComponents(false);
                if (raw is object[] arr) rawComponents = arr;
            }
            catch (Exception ex)
            {
                output.WriteLine($"WARN: GetComponents threw: {ex.Message}");
            }

            var sb = new StringBuilder();
            sb.Append('{');
            sb.Append("\"path\":")  .Append(SwDocSession.JsonString(resolvedPath)); sb.Append(',');
            sb.Append("\"config\":").Append(SwDocSession.JsonString(activeConfig)); sb.Append(',');
            sb.Append("\"components\":[");
            int emitted = 0;
            foreach (var obj in rawComponents)
            {
                // Hard-cast inside try/catch per the
                // as-on-COM-RCW-unreliability lesson. `as IComponent2`
                // would silently return null even when the RCW does
                // implement the interface.
                IComponent2? comp = null;
                try { comp = (IComponent2)obj; } catch { comp = null; }
                if (comp is null) continue;

                string cName = "", cPath = "", cConfig = "";
                bool cSuppressed = false;
                try { cName       = comp.Name2 ?? ""; }              catch { }
                try { cPath       = comp.GetPathName() ?? ""; }      catch { }
                try { cConfig     = comp.ReferencedConfiguration ?? ""; } catch { }
                try { cSuppressed = comp.IsSuppressed(); }           catch { }

                if (emitted > 0) sb.Append(',');
                sb.Append('{');
                sb.Append("\"name\":")      .Append(SwDocSession.JsonString(cName));   sb.Append(',');
                sb.Append("\"path\":")      .Append(SwDocSession.JsonString(cPath));   sb.Append(',');
                sb.Append("\"config\":")    .Append(SwDocSession.JsonString(cConfig)); sb.Append(',');
                sb.Append("\"suppressed\":").Append(cSuppressed ? "true" : "false");
                sb.Append('}');
                emitted++;
            }
            sb.Append("]}");
            output.WriteLine(sb.ToString());
            return Task.FromResult(0);
        }
        finally
        {
            SwDocSession.CloseIfWeOpened(app, doc, weOpenedIt);
        }
    }
}
