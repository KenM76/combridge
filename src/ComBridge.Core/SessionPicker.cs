using System.Runtime.InteropServices;

namespace ComBridge.Core;

/// <summary>
/// Discovers running instances for a plugin and resolves a <c>--session</c>
/// selector ("1", "pid:12345", or a title substring) to one of them.
/// </summary>
public static class SessionPicker
{
    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);

    [DllImport("user32.dll")]
    private static extern IntPtr GetTopWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

    private const uint GW_HWNDNEXT = 2;

    /// <summary>Convert an HWND to its owning process ID via Win32, or null on failure.</summary>
    public static int? PidFromHwnd(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero) return null;
        if (GetWindowThreadProcessId(hwnd, out var pid) == 0) return null;
        return (int)pid;
    }

    /// <summary>
    /// Rank the given PIDs by their top-level windows' position in the
    /// desktop-global Z-order. Lower rank = more recently focused. PIDs whose
    /// processes have no top-level window currently visible in the Z-order
    /// get <see cref="int.MaxValue"/>, sorting them to the end as a fallback.
    /// </summary>
    /// <remarks>
    /// Windows maintains the Z-order globally and updates it whenever any
    /// window gains focus. After the user clicks SW window B, then switches
    /// to a terminal to run combridge, the Z-order is [terminal, B, A, ...]
    /// from top down. The terminal's PID won't be in <paramref name="pids"/>,
    /// so B gets the lowest rank and wins as "most recently used."
    /// </remarks>
    private static Dictionary<int, int> RankByZOrder(IEnumerable<int> pids)
    {
        var wanted = new HashSet<int>(pids);
        var ranks  = new Dictionary<int, int>();
        if (wanted.Count == 0) return ranks;

        int z = 0;
        for (var h = GetTopWindow(IntPtr.Zero); h != IntPtr.Zero; h = GetWindow(h, GW_HWNDNEXT))
        {
            GetWindowThreadProcessId(h, out var pid);
            var p = (int)pid;
            if (wanted.Contains(p) && !ranks.ContainsKey(p))
            {
                ranks[p] = z;
                if (ranks.Count == wanted.Count) break;   // got them all
            }
            z++;
        }
        // PIDs with no matching top-level window (minimized to tray, headless
        // service host, etc.) get a sentinel rank so they sort to the end.
        foreach (var p in wanted)
            if (!ranks.ContainsKey(p)) ranks[p] = int.MaxValue;

        return ranks;
    }

    /// <summary>
    /// Discover every running instance of the app this plugin targets and
    /// return them as a numbered list ready for <c>list-sessions</c> display
    /// or for <see cref="Resolve"/>-based selection.
    /// </summary>
    /// <param name="plugin">
    /// The plugin whose <see cref="IComBridgePlugin.RotMonikerPatterns"/>,
    /// <see cref="IComBridgePlugin.ProgIds"/>, and
    /// <see cref="IComBridgePlugin.TryExtractRoot"/> + <see cref="IComBridgePlugin.DescribeInstance"/>
    /// drive discovery and labeling.
    /// </param>
    /// <returns>
    /// List of <c>(Root, Info)</c> tuples where <c>Root</c> is the connected
    /// COM Application object and <c>Info</c> describes it with a 1-based
    /// <see cref="SessionInfo.Index"/>, PID, title, and display string. The
    /// list is deduplicated by PID and filtered to exclude dead-binding
    /// signatures (entries where both PID and title are missing).
    /// </returns>
    /// <remarks>
    /// Combines two discovery paths and dedupes by PID:
    /// <list type="number">
    ///   <item>ROT-moniker pattern walk via <see cref="RotHelper.EnumerateActiveObjects"/>,
    ///         then <see cref="IComBridgePlugin.TryExtractRoot"/> to ascend
    ///         from per-document monikers (Excel <c>.xlsx</c>, Word <c>.docx</c>)
    ///         to the parent Application. Identity transform for plugins
    ///         whose patterns already match the Application itself (SolidWorks).</item>
    ///   <item><c>oleaut32!GetActiveObject(progId)</c> via
    ///         <see cref="RotHelper.TryCoGetActiveObject"/> — handles apps
    ///         reachable via the legacy Marshal.GetActiveObject mechanism that
    ///         aren't discoverable through plain ROT enumeration (Outlook, and
    ///         the Excel-with-no-saved-workbook case).</item>
    /// </list>
    /// Dead-binding filter: drops entries where <see cref="IComBridgePlugin.DescribeInstance"/>
    /// returns both <c>null</c> PID AND empty title. This catches transient
    /// sidecar processes (notably Office 365's shared-instance shim leaves
    /// hollow EXCEL.EXE processes briefly visible in the ROT before workbook
    /// state migrates back to the existing instance).
    /// <para>
    /// <b>Ordering:</b> the returned list is sorted by <b>desktop Z-order
    /// (most-recently-used first)</b> using <see cref="RankByZOrder"/>.
    /// <see cref="SessionInfo.Index"/> is reassigned 1-based per the
    /// MRU-sorted order, so <c>--session 1</c> always picks the window the
    /// user was most recently focused on, <c>--session 2</c> the next-most,
    /// and so on. Sessions whose process has no visible top-level window
    /// fall to the end (preserving ROT enumeration order among them).
    /// </para>
    /// </remarks>
    public static List<(object Root, SessionInfo Info)> Enumerate(IComBridgePlugin plugin)
    {
        // Two discovery paths, combined and PID-deduped:
        //   1. ROT-moniker pattern walk, then plugin.TryExtractRoot to ascend from
        //      per-doc monikers (Excel .xlsx) to the parent application. Identity for
        //      plugins whose patterns already match the application (SolidWorks).
        //   2. oleaut32!GetActiveObject — single instance, no TryExtractRoot needed
        //      because the returned object is already the application root. Handles
        //      apps reachable via Marshal.GetActiveObject-style lookup (Excel with
        //      no workbook open, or plugins that don't define ROT patterns).
        var roots = new List<object>();
        foreach (var raw in RotHelper.EnumerateActiveObjects(plugin.RotMonikerPatterns))
        {
            object? root = null;
            try { root = plugin.TryExtractRoot(raw); } catch { /* skip broken bindings */ }
            if (root is not null) roots.Add(root);
        }
        foreach (var progId in plugin.ProgIds)
        {
            var obj = RotHelper.TryCoGetActiveObject(progId);
            if (obj is not null) roots.Add(obj);
        }
        var seenPids = new HashSet<int>();
        // Discovery pass: build (root, partial SessionInfo) tuples in ROT-walk
        // order. Index is assigned after MRU sorting below.
        var discovered = new List<(object Root, int? Pid, string? Title, string Desc)>();
        foreach (var root in roots)
        {
            (int? pid, string? title) = (null, null);
            try { (pid, title) = plugin.DescribeInstance(root); } catch { /* tolerate broken instances */ }

            // Drop zombie / transient sidecar processes. Some COM hosts publish
            // short-lived secondary processes that briefly own a moniker, then
            // lose their main window (Hwnd→0, so PidFromHwnd→null) and their
            // active document (title→null) within seconds. The underlying RCW
            // is then unusable but still enumerable in the ROT. Office 365's
            // shared-instance shim is the prototypical example (a new EXCEL.EXE
            // spawned via CoCreateInstance briefly hosts a workbook before
            // Office migrates state back to the existing instance), but the
            // pattern can show up with any COM server that uses helper
            // processes or hand-off protocols.
            //
            // Rule: if the plugin's DescribeInstance returns BOTH null pid AND
            // empty title, treat the binding as dead. Keep entries with at
            // least one usable signal — e.g. a real SW instance with no model
            // loaded legitimately has title=null but a valid pid.
            if (pid is null && string.IsNullOrEmpty(title)) continue;

            // Dedupe by PID — SW and Excel both register multiple monikers per process.
            if (pid is int p && !seenPids.Add(p)) continue;

            var desc = (pid, title) switch
            {
                (int pp, string t) when !string.IsNullOrEmpty(t) => $"pid={pp}  title={t}",
                (int pp, _)                                       => $"pid={pp}",
                (null, string t) when !string.IsNullOrEmpty(t)    => t,
                _                                                  => "(no info)",
            };
            discovered.Add((root, pid, title, desc));
        }

        // Sort by desktop Z-order: most-recently-focused window first.
        // PIDs with no Z-order entry (null pid, or process with no top-level
        // window) sort to the end, preserving discovery order among themselves.
        var ranks = RankByZOrder(discovered.Where(d => d.Pid is not null).Select(d => d.Pid!.Value));
        var orderedDiscovery = discovered
            .Select((d, originalIdx) => (d, originalIdx))
            .OrderBy(x => x.d.Pid is int p && ranks.TryGetValue(p, out var r) ? r : int.MaxValue)
            .ThenBy(x => x.originalIdx)
            .Select(x => x.d)
            .ToList();

        // Reassign 1-based Index in the MRU-sorted order so list-sessions and
        // --session N agree with what the user sees.
        var sessions = new List<(object, SessionInfo)>();
        int newIdx = 1;
        foreach (var (root, pid, title, desc) in orderedDiscovery)
            sessions.Add((root, new SessionInfo(newIdx++, pid, title, desc)));

        return sessions;
    }

    /// <summary>
    /// Resolve a <c>--session</c> selector against the enumerated list.
    /// </summary>
    /// <remarks>
    /// Selector forms (case-insensitive where alphabetic):
    /// <list type="bullet">
    ///   <item>null / empty — the most-recently-focused session (the default;
    ///         <see cref="Enumerate"/> sorts MRU-first).</item>
    ///   <item><c>last</c> / <c>mru</c> / <c>recent</c> — explicit MRU keyword.
    ///         Equivalent to "no selector"; useful when a script wants to be
    ///         explicit and defensive against future default changes.</item>
    ///   <item>pure digits, e.g. <c>2</c> — match by 1-based
    ///         <see cref="SessionInfo.Index"/> (MRU order, so <c>1</c> is most
    ///         recent, <c>2</c> next-most, etc.).</item>
    ///   <item><c>pid:NNNN</c> — match by Win32 process ID.</item>
    ///   <item>any other string — case-insensitive substring of the session's
    ///         title (or the full description string).</item>
    /// </list>
    /// </remarks>
    public static (object Root, SessionInfo Info)? Resolve(
        List<(object Root, SessionInfo Info)> sessions, string? selector)
    {
        if (sessions.Count == 0) return null;
        if (string.IsNullOrEmpty(selector)) return sessions[0];

        // Explicit MRU keyword: same result as "no selector," but lets scripts
        // be defensive against future default changes.
        if (string.Equals(selector, "last",   StringComparison.OrdinalIgnoreCase) ||
            string.Equals(selector, "mru",    StringComparison.OrdinalIgnoreCase) ||
            string.Equals(selector, "recent", StringComparison.OrdinalIgnoreCase))
            return sessions[0];

        if (int.TryParse(selector, out var n))
            return sessions.FirstOrDefault(s => s.Info.Index == n) is { } match0
                ? match0 : null;

        if (selector.StartsWith("pid:", StringComparison.OrdinalIgnoreCase) &&
            int.TryParse(selector.AsSpan(4), out var pid))
            return sessions.FirstOrDefault(s => s.Info.Pid == pid) is { } match1
                ? match1 : null;

        return sessions.FirstOrDefault(s =>
            (s.Info.Title?.Contains(selector, StringComparison.OrdinalIgnoreCase) ?? false) ||
            s.Info.Description.Contains(selector, StringComparison.OrdinalIgnoreCase)) is { } match2
            ? match2 : null;
    }
}
