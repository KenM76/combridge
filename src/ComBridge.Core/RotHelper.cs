using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.Versioning;
using System.Text.RegularExpressions;

namespace ComBridge.Core;

/// <summary>
/// Attaches to an already-running COM server via the Running Object Table (ROT).
/// In .NET (Core/5+) <c>Marshal.GetActiveObject</c> was removed, so we re-implement
/// it via the same OLE32 P/Invokes the BCL used to call.
/// </summary>
[SupportedOSPlatform("windows")]
public static class RotHelper
{
    [DllImport("ole32.dll")]
    private static extern int GetRunningObjectTable(uint reserved, out IRunningObjectTable prot);

    [DllImport("ole32.dll")]
    private static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);

    [DllImport("ole32.dll", PreserveSig = true)]
    private static extern int CLSIDFromProgID(
        [MarshalAs(UnmanagedType.LPWStr)] string lpszProgId,
        out Guid pclsid);

    /// <summary>
    /// <c>oleaut32!GetActiveObject</c>. This is the underlying API that
    /// <c>Marshal.GetActiveObject</c> (removed from modern .NET) used to call.
    /// Unlike walking the ROT and matching moniker display names, this asks
    /// COM's own resolution machinery — which knows how to locate apps (like
    /// Excel) that don't register a class moniker in the ROT.
    /// </summary>
    [DllImport("oleaut32.dll", PreserveSig = false)]
    [return: MarshalAs(UnmanagedType.IUnknown)]
    private static extern object GetActiveObject([In] ref Guid rclsid, IntPtr reserved);

    /// <summary>
    /// Attach to a running instance via <c>oleaut32!GetActiveObject(CLSID)</c>.
    /// Returns the single "default" running instance (usually first registered)
    /// or null. Use this when a server doesn't register a discoverable moniker
    /// in the ROT — most notably Excel, which registers only with interface
    /// monikers that aren't directly QI-able to <c>_Application</c>.
    /// </summary>
    public static object? TryCoGetActiveObject(string progId)
    {
        if (CLSIDFromProgID(progId, out var clsid) != 0) return null;
        try { return GetActiveObject(ref clsid, IntPtr.Zero); }
        catch (COMException) { return null; }   // MK_E_UNAVAILABLE etc.
    }

    /// <summary>
    /// Try to attach to a running instance whose ROT moniker matches the given
    /// regex pattern. Returns null if no match.
    /// </summary>
    public static object? TryGetActiveObject(string rotPattern)
        => EnumerateActiveObjects(new[] { rotPattern }).FirstOrDefault();

    /// <summary>
    /// Enumerate ALL running instances whose ROT display name matches any of the
    /// given regex patterns (case-insensitive). Order is the order returned by
    /// <c>IEnumMoniker</c>, which is approximately registration (launch) order
    /// but not strictly guaranteed. Caller is responsible for deduping (e.g. by
    /// PID) — a single process may register multiple monikers.
    /// </summary>
    /// <param name="rotPatterns">
    /// .NET regex patterns. Each is compiled once and matched case-insensitively
    /// against each ROT entry's <c>IMoniker.GetDisplayName</c>. See
    /// <see cref="IComBridgePlugin.RotMonikerPatterns"/> for real-world examples.
    /// </param>
    public static IEnumerable<object> EnumerateActiveObjects(IEnumerable<string> rotPatterns)
    {
        var regexes = rotPatterns
            .Select(p => new Regex(p, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            .ToArray();
        if (regexes.Length == 0) yield break;
        if (GetRunningObjectTable(0, out var rot) != 0 || rot is null) yield break;
        if (CreateBindCtx(0, out var bindCtx) != 0 || bindCtx is null) yield break;

        rot.EnumRunning(out var monikerEnum);
        monikerEnum.Reset();
        var moniker = new IMoniker[1];
        var fetched = IntPtr.Zero;

        while (monikerEnum.Next(1, moniker, fetched) == 0)
        {
            moniker[0].GetDisplayName(bindCtx, null!, out var displayName);
            if (displayName is null) continue;
            if (!regexes.Any(r => r.IsMatch(displayName))) continue;
            if (rot.GetObject(moniker[0], out var obj) == 0 && obj is not null)
                yield return obj;
        }
    }

    /// <summary>
    /// Attach to a running instance via ROT (matching any moniker pattern), or
    /// create a new one if <paramref name="createIfMissing"/>. Throws if neither
    /// attach nor create succeeds.
    /// </summary>
    /// <param name="rotPatterns">Regex patterns for ROT moniker matching (attach path).</param>
    /// <param name="progIds">ProgIDs used for <c>CreateInstance</c> when attach fails.</param>
    /// <param name="createIfMissing">If true and attach fails, instantiate via the first ProgID that resolves.</param>
    public static object AttachOrCreate(IEnumerable<string> rotPatterns, IEnumerable<string> progIds, bool createIfMissing)
        => AttachOrCreate(rotPatterns, progIds, transformer: null, createIfMissing);

    /// <summary>
    /// Plugin-aware overload. Calls <paramref name="transformer"/> on each
    /// ROT-bound object before returning (e.g. Excel's Workbook→Application
    /// ascent). <c>null</c> from the transformer skips that binding.
    /// </summary>
    public static object AttachOrCreate(
        IEnumerable<string> rotPatterns,
        IEnumerable<string> progIds,
        Func<object, object?>? transformer,
        bool createIfMissing)
    {
        var patterns = rotPatterns.ToArray();
        var ids = progIds.ToArray();

        // 1. ROT-moniker pattern walk + optional transform (multi-instance servers).
        foreach (var raw in EnumerateActiveObjects(patterns))
        {
            object? root = transformer is null ? raw : transformer(raw);
            if (root is not null) return root;
        }

        // 2. GetActiveObject fallback (single-instance servers, Excel-style).
        foreach (var id in ids)
        {
            var obj = TryCoGetActiveObject(id);
            if (obj is not null) return obj;
        }

        if (!createIfMissing)
            throw new InvalidOperationException(
                $"No running instance found via ROT patterns ({string.Join(", ", patterns)}) " +
                $"or GetActiveObject ({string.Join(", ", ids)}).");

        foreach (var id in ids)
        {
            var t = Type.GetTypeFromProgID(id, throwOnError: false);
            if (t is not null)
            {
                var inst = Activator.CreateInstance(t);
                if (inst is not null) return inst;
            }
        }
        throw new InvalidOperationException(
            $"Could not attach (patterns: {string.Join(", ", patterns)}) or create (ProgIDs: {string.Join(", ", ids)}).");
    }

    /// <summary>
    /// Backward-compatible overload: treat ProgIDs both as attach patterns and as
    /// creation ProgIDs. Library-mode consumers calling
    /// <c>RotHelper.AttachOrCreate(new[] { "SwDocumentMgr.SwDocumentMgr" }, true)</c>
    /// keep working — DocMgr-style stateless servers are usually created, not
    /// attached, and their ProgIDs work as patterns when escaped.
    /// </summary>
    public static object AttachOrCreate(IEnumerable<string> progIds, bool createIfMissing)
    {
        var ids = progIds.ToArray();
        var patterns = ids.Select(Regex.Escape).ToArray();
        return AttachOrCreate(patterns, ids, createIfMissing);
    }
}
