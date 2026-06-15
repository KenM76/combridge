using ComBridge.Core;
using Microsoft.CodeAnalysis;
using Xl = global::Microsoft.Office.Interop.Excel;

namespace ComBridge.Plugins.Excel;

/// <summary>
/// Globals exposed to user .csx scripts. Each is null-safe — Excel may be open
/// with no workbook, or with a workbook but the active sheet might not be a
/// Worksheet (it could be a Chart). Scripts should null-check before use.
/// <para>
/// <c>xlApp</c> is typed as <see cref="Xl._Application"/> (the dispinterface),
/// not <c>Xl.Application</c> (the co-class), because the COM RCW returned by
/// <c>IRunningObjectTable.GetObject</c> for Excel's moniker only supports the
/// dispinterface IID. A direct cast to the co-class fails on ROT-attached
/// instances. The dispinterface has the same member surface scripts care about.
/// </para>
/// </summary>
public sealed class ExcelGlobals : IScriptContext
{
    public Xl._Application xlApp { get; }
    public Xl.Workbook? xlBook { get; }
    public Xl.Worksheet? xlSheet { get; }

    // Host-injected I/O channels — see IScriptContext.
    public string[] ScriptArgs { get; set; } = Array.Empty<string>();
    public string Stdin { get; set; } = "";

    internal ExcelGlobals(Xl._Application app)
    {
        xlApp = app;
        try { xlBook = app.ActiveWorkbook; } catch { xlBook = null; }
        try { xlSheet = app.ActiveSheet as Xl.Worksheet; } catch { xlSheet = null; }
    }
}

public sealed class ExcelPlugin : IComBridgePlugin
{
    public string Name => "excel";
    public string Description => "Microsoft Excel (attach via ROT or launch new instance). Globals: xlApp, xlBook, xlSheet.";

    // Excel.Application is the canonical ProgID for both ROT attach and
    // Activator.CreateInstance. ROT registration is conditional: Excel
    // registers when UserControl=true (the default after interactive launch
    // with a workbook open). When absent, the host falls back to creating a
    // new instance because AllowCreateNew is true.
    public string[] ProgIds => new[] { "Excel.Application" };
    public bool AllowCreateNew => true;
    public Type GlobalsType => typeof(ExcelGlobals);

    // Excel does NOT publish its _Application interface as a discoverable
    // moniker the way SolidWorks does. What IS in the ROT are per-workbook
    // file monikers (one per open .xlsx / .xlsm / .xlsb / template / add-in).
    // For multi-instance discovery: match file extensions; each binding gives
    // a Workbook RCW; TryExtractRoot ascends Workbook.Application to get the
    // parent Excel app. PID dedupe in SessionPicker collapses multiple
    // workbooks in the same process to one session.
    //
    // SessionPicker ALSO calls TryCoGetActiveObject(ProgIds) — that handles
    // the rare "Excel running with no workbook open" case, and acts as a
    // belt-and-braces for the most-recent active instance.
    public IEnumerable<string> RotMonikerPatterns => new[]
    {
        @"\.(xlsx|xls|xlsm|xlsb|xltx|xltm|xla|xlam)$",
    };

    /// <summary>
    /// Ascend from a file-moniker binding (a Workbook / AddIn RCW) to its
    /// parent Excel Application. The bound object's <c>.Application</c>
    /// property is an IDispatch — dynamic dispatch handles it without the
    /// dispinterface-vs-coclass cast quirk that direct interop would hit.
    /// </summary>
    public object? TryExtractRoot(object monikerBound)
    {
        try
        {
            dynamic d = monikerBound;
            object? app = d.Application;
            return app;   // SessionPicker / CreateGlobals will cast to Xl._Application
        }
        catch
        {
            return null;   // unexpected binding (not a workbook/addin), skip
        }
    }

    public object CreateGlobals(object comRoot) => new ExcelGlobals((Xl._Application)comRoot);

    public IEnumerable<MetadataReference> ScriptReferences
    {
        get
        {
            var here = Path.GetDirectoryName(typeof(ExcelPlugin).Assembly.Location)!;
            // Office PIAs deployed alongside the plugin: Microsoft.Office.Interop.Xl.dll,
            // office.dll, Microsoft.Vbe.Interop.dll, etc.
            foreach (var dll in Directory.EnumerateFiles(here, "Microsoft.Office.*.dll"))
                yield return MetadataReference.CreateFromFile(dll);
            foreach (var dll in Directory.EnumerateFiles(here, "office.dll"))
                yield return MetadataReference.CreateFromFile(dll);
            foreach (var dll in Directory.EnumerateFiles(here, "Microsoft.Vbe.*.dll"))
                yield return MetadataReference.CreateFromFile(dll);
        }
    }

    public IEnumerable<string> ScriptUsings => new[]
    {
        "Microsoft.Office.Interop.Excel",
    };

    public IEnumerable<IBridgeCommand> Commands => new IBridgeCommand[]
    {
        new InfoCommand(),
        new DumpSheetCommand(),
        new NewScriptCommand(),
        new ListAddinsCommand(),
    };

    /// <summary>
    /// Excel exposes its main-window HWND as <c>Application.Hwnd</c>. Combined with
    /// <c>ActiveWorkbook.Name</c> this gives the session picker a clean PID + title.
    /// </summary>
    public (int? Pid, string? Title) DescribeInstance(object comRoot)
    {
        try
        {
            var app = (Xl._Application)comRoot;
            int? pid = null;
            try
            {
                var hwnd = (IntPtr)app.Hwnd;
                pid = SessionPicker.PidFromHwnd(hwnd);
            }
            catch { }

            string? title = null;
            try { title = app.ActiveWorkbook?.Name; } catch { }
            return (pid, title);
        }
        catch
        {
            return (null, null);
        }
    }
}

/// <summary>Smoke-test: prints Excel version and active workbook/sheet names.</summary>
internal sealed class InfoCommand : IBridgeCommand
{
    public string Name => "info";
    public string Usage => "info   (prints Excel version + active workbook/sheet)";

    public Task<int> RunAsync(object comRoot, string[] args, TextWriter output)
    {
        var app = (Xl._Application)comRoot;
        output.WriteLine($"Excel version: {app.Version}");
        output.WriteLine($"Visible:       {app.Visible}");
        try { output.WriteLine($"ActiveWorkbook: {app.ActiveWorkbook?.Name ?? "(none)"}"); }
        catch { output.WriteLine("ActiveWorkbook: (unavailable)"); }
        try
        {
            var sheet = app.ActiveSheet as Xl.Worksheet;
            output.WriteLine($"ActiveSheet:    {sheet?.Name ?? "(none / not a Worksheet)"}");
        }
        catch { output.WriteLine("ActiveSheet:    (unavailable)"); }
        return Task.FromResult(0);
    }
}

/// <summary>
/// <c>dump-sheet [sheetName]</c> — dump the used range of the active (or named) sheet
/// as tab-separated rows. Useful for quickly grabbing data without writing a script.
/// </summary>
internal sealed class DumpSheetCommand : IBridgeCommand
{
    public string Name => "dump-sheet";
    public string Usage => "dump-sheet [sheetName]   (dumps used range as TSV)";

    public Task<int> RunAsync(object comRoot, string[] args, TextWriter output)
    {
        var app = (Xl._Application)comRoot;
        var book = app.ActiveWorkbook ?? throw new InvalidOperationException("No active workbook.");
        Xl.Worksheet sheet;
        if (args.Length >= 1)
            sheet = (Xl.Worksheet)book.Worksheets[args[0]];
        else
            sheet = (Xl.Worksheet)(app.ActiveSheet
                ?? throw new InvalidOperationException("No active sheet."));

        var used = sheet.UsedRange;
        // Range.Value returns object[,] for multi-cell ranges, scalar for single cell.
        // 1-based indexing.
        var raw = used.Value;
        if (raw is object[,] arr)
        {
            int rows = arr.GetLength(0);
            int cols = arr.GetLength(1);
            for (int r = 1; r <= rows; r++)
            {
                for (int c = 1; c <= cols; c++)
                {
                    if (c > 1) output.Write('\t');
                    output.Write(arr[r, c]?.ToString() ?? "");
                }
                output.WriteLine();
            }
        }
        else
        {
            output.WriteLine(raw?.ToString() ?? "");
        }
        return Task.FromResult(0);
    }
}

/// <summary>
/// <c>excel new-script &lt;path&gt; [--force]</c> — scaffolds a starter <c>.csx</c>
/// at <c>&lt;path&gt;</c> showing the conventional Excel-interop alias pattern,
/// the available globals, and a minimal example body the author edits.
/// <para>
/// Replaces the v0.4.2 host-side preamble mechanism (ScriptUsingAliases) that
/// was withdrawn in v0.5.0. The scaffold gives authors the same zero-typing
/// onboarding while preserving the source-is-truth contract: the <c>.csx</c>
/// the host compiles is exactly the file the author (and any IDE, LLM, or
/// app-store auditor) can read.
/// </para>
/// </summary>
internal sealed class NewScriptCommand : IBridgeCommand
{
    public string Name => "new-script";
    public string Usage => "new-script <path> [--force]   (scaffolds a starter .csx with the Xl alias + a minimal example)";

    private const string Template =
@"// <NAME> — Excel script for combridge.
//
// Globals provided by ComBridge.Plugins.Excel:
//   xlApp   : Excel._Application — the running Excel instance
//   xlBook  : Excel.Workbook?    — active workbook (null if none open)
//   xlSheet : Excel.Worksheet?   — active worksheet (null if a Chart is active)
//
// Office interop namespaces collide with several BCL types (Range, Exception,
// Application, Style, Font, Action, Page). The recommended convention is to
// alias the interop namespace to a two-letter name (`Xl` for Excel), use the
// qualified form for Office types, and fully-qualify the BCL side as needed.
// This single line is the entire ergonomic fix — every reader of this script
// (you, your IDE, an LLM, an app-store reviewer) can see exactly what's in
// scope without consulting host docs. See LLM/scripting.md for the full
// collision table.
using Xl = global::Microsoft.Office.Interop.Excel;

if (xlBook is null) {
    Console.WriteLine(""ERROR: no workbook open in Excel."");
    return 1;
}

Console.WriteLine($""Active workbook: {xlBook.Name}"");
Console.WriteLine($""Active sheet:    {xlSheet?.Name ?? ""(none)""}"");

// Example: report the dimensions of the used range on the active sheet.
Xl.Range used = xlSheet?.UsedRange;
if (used is not null) {
    Console.WriteLine($""Used range:  {used.Rows.Count} rows x {used.Columns.Count} cols"");
}

return 0;
";

    public Task<int> RunAsync(object comRoot, string[] args, TextWriter output)
        => ScriptScaffold.WriteTemplate(args, output, Template, Name, Usage);
}
