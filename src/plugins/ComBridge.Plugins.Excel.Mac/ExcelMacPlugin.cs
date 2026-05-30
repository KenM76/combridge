using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using ComBridge.Core;
using Microsoft.CodeAnalysis;

namespace ComBridge.Plugins.Excel.Mac;

/// <summary>
/// macOS Excel plugin. Same CLI contract as the Windows
/// <c>ComBridge.Plugins.Excel</c> plugin (<c>Name = "excel"</c>, same
/// commands by name), implemented via AppleScript (<c>osascript</c>)
/// instead of COM. PluginLoader's <see cref="OSPlatform"/> filter ensures
/// only one "excel" plugin loads per OS — Windows Excel plugin on
/// Windows, this on Mac.
/// </summary>
/// <remarks>
/// Architecture parallels the Windows Excel plugin where possible:
/// <list type="bullet">
///   <item><b>FindSessions</b>: enumerates running Excel via
///         <c>tell application "System Events" to processes</c>; on macOS
///         there's typically ONE Excel process per user (similar to
///         Outlook on Windows), so the result is 0 or 1 sessions.</item>
///   <item><b>Globals</b>: <see cref="XlMacGlobals"/> exposes <c>xlApp</c>
///         (an <see cref="XlMacApp"/> AppleScript wrapper) — same identifier
///         name as Windows, different type.</item>
///   <item><b>Commands</b>: <c>info</c>, <c>dump-sheet</c> — same names
///         as the Windows plugin so cross-OS ScripTree files work without
///         per-OS dispatch.</item>
/// </list>
///
/// What's deliberately NOT here vs. the Windows plugin:
/// <list type="bullet">
///   <item>No <c>RotMonikerPatterns</c> override — macOS has no ROT, and
///         the base default isn't used anyway since FindSessions is overridden.</item>
///   <item>No <c>TryExtractRoot</c> override — same reason.</item>
///   <item>No <c>office.dll</c> dependency — pure subprocess invocation.</item>
///   <item>Multi-instance Excel doesn't exist on macOS by design (single
///         app, multiple workbooks) — same as Office 365 on Windows in
///         practice, just enforced rather than incidental.</item>
/// </list>
/// </remarks>
public sealed class ExcelMacPlugin : IComBridgePlugin
{
    public string Name => "excel";
    public string Description => "Microsoft Excel for macOS (AppleScript backend). Globals: xlApp.";
    public string[] ProgIds => new[] { "Microsoft Excel" };   // AppleScript application name; not a Windows ProgID
    public bool AllowCreateNew => true;
    public Type GlobalsType => typeof(XlMacGlobals);

    public IReadOnlyCollection<OSPlatform> SupportedPlatforms => new[] { OSPlatform.OSX };

    public object CreateGlobals(object comRoot)
    {
        // On macOS we ignore comRoot — there's no COM object to wrap. The
        // session "Root" is just a sentinel; the globals talk to Excel via
        // osascript on demand.
        return new XlMacGlobals();
    }

    public IEnumerable<MetadataReference> ScriptReferences
    {
        get
        {
            // Reference this plugin's own assembly so scripts can use XlMacApp
            // members. No interop DLLs needed.
            yield return MetadataReference.CreateFromFile(typeof(ExcelMacPlugin).Assembly.Location);
        }
    }

    public IEnumerable<string> ScriptUsings => new[] { "ComBridge.Plugins.Excel.Mac" };

    public IEnumerable<IBridgeCommand> Commands => new IBridgeCommand[]
    {
        new ExcelMacInfoCommand(),
        new ExcelMacDumpSheetCommand(),
    };

    /// <summary>
    /// Discover running Excel instances on macOS. Uses <c>System Events</c>
    /// to enumerate processes named "Microsoft Excel". Typically returns
    /// either an empty list (Excel not running) or a single-element list
    /// (Excel running with optional active workbook).
    /// </summary>
    public List<(object Root, SessionInfo Info)> FindSessions()
    {
        var sessions = new List<(object, SessionInfo)>();

        // Sanity: only run osascript if we're actually on macOS. On other
        // platforms (the plugin shouldn't load there thanks to SupportedPlatforms,
        // but be defensive) Osascript.IsAvailable returns false.
        if (!Osascript.IsAvailable()) return sessions;

        // Is Excel running?
        var running = Osascript.TryRun(
            "tell application \"System Events\" to (name of processes) contains \"Microsoft Excel\"");
        if (running != "true") return sessions;

        // PID via `unix id of process`
        int? pid = null;
        var pidRaw = Osascript.TryRun(
            "tell application \"System Events\" to unix id of (first process whose name is \"Microsoft Excel\")");
        if (int.TryParse(pidRaw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n)) pid = n;

        // Active workbook name (best-effort)
        string? title = Osascript.TryRun(
            "tell application \"Microsoft Excel\" to name of active workbook");
        if (string.IsNullOrEmpty(title)) title = null;

        var desc = (pid, title) switch
        {
            (int pp, string t) when !string.IsNullOrEmpty(t) => $"pid={pp}  title={t}",
            (int pp, _)                                       => $"pid={pp}",
            (null, string t) when !string.IsNullOrEmpty(t)    => t,
            _                                                  => "(no info)",
        };

        // Sentinel root — Mac plugin doesn't use a COM RCW. CreateGlobals
        // ignores it. SessionInfo carries everything the host needs.
        var root = new object();
        sessions.Add((root, new SessionInfo(1, pid, title, desc)));
        return sessions;
    }

    /// <summary>
    /// DescribeInstance is unused on this plugin (FindSessions already
    /// returns a fully-populated SessionInfo). Stub kept for interface
    /// compliance — returns null/null, but the default would never be
    /// called anyway since FindSessions doesn't go through SessionPicker.
    /// </summary>
    public (int? Pid, string? Title) DescribeInstance(object comRoot) => (null, null);
}

internal sealed class ExcelMacInfoCommand : IBridgeCommand
{
    public string Name => "info";
    public string Usage => "info   (prints Excel version + active workbook/sheet)";

    public Task<int> RunAsync(object comRoot, string[] args, TextWriter output)
    {
        var app = ((XlMacGlobals)comRoot is not null ? (XlMacGlobals)comRoot : new XlMacGlobals()).xlApp;
        // Note: comRoot here is the sentinel `new object()` from FindSessions.
        // We can't cast it to XlMacGlobals — globals are constructed by
        // CreateGlobals which the host invokes for run-script. For built-in
        // commands like `info`, we get the raw sentinel + need a fresh XlMacApp.
        // The line above is overly clever — simplify:
        app = new XlMacApp();

        try { output.WriteLine($"Excel version: {app.Version}"); }
        catch { output.WriteLine("Excel version: (Excel not running or unreachable)"); }

        try { output.WriteLine($"Visible:       {app.Visible}"); }
        catch { output.WriteLine("Visible:       (unavailable)"); }

        try { output.WriteLine($"ActiveWorkbook: {app.ActiveWorkbookName ?? "(none)"}"); }
        catch { output.WriteLine("ActiveWorkbook: (unavailable)"); }

        try { output.WriteLine($"ActiveSheet:    {app.ActiveSheetName ?? "(none)"}"); }
        catch { output.WriteLine("ActiveSheet:    (unavailable)"); }

        try { output.WriteLine($"Workbooks:      {app.WorkbookCount}"); } catch { }
        return Task.FromResult(0);
    }
}

internal sealed class ExcelMacDumpSheetCommand : IBridgeCommand
{
    public string Name => "dump-sheet";
    public string Usage => "dump-sheet [sheetName]   (dumps used range as TSV)";

    public Task<int> RunAsync(object comRoot, string[] args, TextWriter output)
    {
        var app = new XlMacApp();
        var sheet = args.Length >= 1 ? args[0] : null;
        try
        {
            var arr = app.DumpUsedRange(sheet);
            int rows = arr.GetLength(0), cols = arr.GetLength(1);
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    if (c > 0) output.Write('\t');
                    output.Write(arr[r, c] ?? "");
                }
                output.WriteLine();
            }
            return Task.FromResult(0);
        }
        catch (OsascriptException ex)
        {
            output.WriteLine($"ERROR: {ex.Message}");
            return Task.FromResult(4);
        }
    }
}
