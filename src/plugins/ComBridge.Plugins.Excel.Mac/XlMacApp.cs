using System.Globalization;

namespace ComBridge.Plugins.Excel.Mac;

/// <summary>
/// Mac Excel application wrapper exposed to user scripts as <c>xlApp</c>.
/// Mirrors a subset of <c>Microsoft.Office.Interop.Excel._Application</c>
/// — version, visibility, active workbook, workbook enumeration — but
/// implemented by shelling out to <c>osascript</c> instead of via COM.
/// </summary>
/// <remarks>
/// <para>
/// Surface parity with the Windows plugin: methods/properties named here
/// match what the Windows ExcelGlobals exposes as <c>xlApp</c>, so a
/// well-written cross-OS .csx can use the same idiom on both sides — but
/// see <c>LLM/extending.md</c> § Mac caveats: Windows-typed members like
/// <c>app.Workbooks[1].Worksheets[2].Cells[3,4].Value</c> won't compile
/// here because the types are different.
/// </para>
/// <para>
/// Application name handling: AppleScript targets Excel as
/// <c>"Microsoft Excel"</c> (NOT <c>"Excel"</c>). This is hardcoded as a
/// constant; if a future Office for Mac rename happens we'd update it
/// here.
/// </para>
/// </remarks>
public sealed class XlMacApp
{
    private const string AppName = "Microsoft Excel";

    /// <summary>Excel version string, e.g. "16.85".</summary>
    public string Version =>
        Osascript.Run($"tell application \"{AppName}\" to version");

    /// <summary>Whether Excel's UI is visible. Mac Excel is almost always visible.</summary>
    public bool Visible
    {
        get
        {
            var raw = Osascript.TryRun($"tell application \"System Events\" to visible of (first process whose name is \"{AppName}\")");
            return raw == "true";
        }
    }

    /// <summary>
    /// Number of currently open workbooks. Returns 0 if Excel isn't running
    /// (rather than throwing — matches the Windows plugin's tolerance).
    /// </summary>
    public int WorkbookCount
    {
        get
        {
            var raw = Osascript.TryRun($"tell application \"{AppName}\" to count of workbooks");
            return int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : 0;
        }
    }

    /// <summary>
    /// Names of every open workbook, in Excel's enumeration order. Empty
    /// array if none.
    /// </summary>
    public string[] WorkbookNames
    {
        get
        {
            // AppleScript returns a comma-space-separated list of names when
            // coerced to text. We split on ", " to recover individual names.
            // Edge case: a workbook name containing ", " literally would be
            // mis-split; rare enough to accept for v0.3.0.
            var raw = Osascript.TryRun(
                $"tell application \"{AppName}\" to return name of every workbook as text");
            if (string.IsNullOrEmpty(raw)) return Array.Empty<string>();
            return raw.Split(", ", StringSplitOptions.RemoveEmptyEntries);
        }
    }

    /// <summary>Name of the active workbook, or null if none open.</summary>
    public string? ActiveWorkbookName
    {
        get
        {
            var raw = Osascript.TryRun($"tell application \"{AppName}\" to name of active workbook");
            return string.IsNullOrEmpty(raw) ? null : raw;
        }
    }

    /// <summary>Name of the active sheet (of the active workbook), or null.</summary>
    public string? ActiveSheetName
    {
        get
        {
            var raw = Osascript.TryRun($"tell application \"{AppName}\" to name of active sheet");
            return string.IsNullOrEmpty(raw) ? null : raw;
        }
    }

    /// <summary>
    /// Get a single cell's value from the active sheet of the active workbook.
    /// Returns null on any failure (no active workbook, cell empty, bad reference).
    /// </summary>
    /// <param name="cellRef">A1-style reference, e.g. <c>"A1"</c>, <c>"C12"</c>.</param>
    public object? GetCellValue(string cellRef)
    {
        var raw = Osascript.TryRun(
            $"tell application \"{AppName}\" to value of cell \"{Osascript.EscapeForAppleScript(cellRef)}\" of active sheet");
        return string.IsNullOrEmpty(raw) ? null : raw;
    }

    /// <summary>
    /// Dump the used range of the active sheet (or the named sheet if non-null)
    /// as a 2D string array. Each call is ONE osascript invocation —
    /// far cheaper than per-cell calls. Returns empty 2D if the workbook has
    /// no used range or Excel isn't running.
    /// </summary>
    public string[,] DumpUsedRange(string? sheetName = null)
    {
        // AppleScript: get value of used range, return as TSV-by-rows newline-separated.
        // The "value of used range" is a list-of-lists; we coerce each row to tab-joined
        // text, then join rows with linefeed.
        var sheetRef = sheetName is null
            ? "active sheet"
            : $"sheet \"{Osascript.EscapeForAppleScript(sheetName)}\" of active workbook";

        var script = $@"
tell application ""{AppName}""
    set theSheet to {sheetRef}
    set theValues to value of used range of theSheet
    set AppleScript's text item delimiters to tab
    set rowLines to {{}}
    repeat with aRow in theValues
        set end of rowLines to (aRow as text)
    end repeat
    set AppleScript's text item delimiters to linefeed
    return rowLines as text
end tell".Trim();

        var raw = Osascript.TryRun(script);
        if (string.IsNullOrEmpty(raw)) return new string[0, 0];

        var rows = raw.Split('\n');
        if (rows.Length == 0) return new string[0, 0];

        // Determine width from the widest row (Excel's used range is rectangular,
        // but be defensive against any AppleScript quirk).
        int maxCols = 0;
        var split = new string[rows.Length][];
        for (int r = 0; r < rows.Length; r++)
        {
            split[r] = rows[r].Split('\t');
            if (split[r].Length > maxCols) maxCols = split[r].Length;
        }

        var arr = new string[rows.Length, maxCols];
        for (int r = 0; r < rows.Length; r++)
            for (int c = 0; c < split[r].Length; c++)
                arr[r, c] = split[r][c];

        return arr;
    }
}

/// <summary>
/// Globals exposed to user .csx scripts on macOS. Single field <c>xlApp</c>
/// wraps the Mac AppleScript surface. Parallels the Windows
/// <c>ExcelGlobals</c> structure but with a different type — scripts that
/// want to work on both OSes must check or branch.
/// </summary>
public sealed class XlMacGlobals
{
    public XlMacApp xlApp { get; }

    internal XlMacGlobals()
    {
        xlApp = new XlMacApp();
    }
}
