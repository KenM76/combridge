using ComBridge.Core;
using ComBridge.Mac.Common;

namespace ComBridge.Plugins.Excel.Mac;

/// <summary>
/// <c>excel list-addins</c> for macOS — best-effort enumeration via
/// AppleScript's <c>addins</c> collection on Microsoft Excel.
/// </summary>
/// <remarks>
/// <para>
/// Mac Excel's AppleScript dictionary exposes <c>addins of application</c>
/// returning a list of <c>add-in</c> records, each with <c>name</c>,
/// <c>full name</c> (path), and <c>installed</c> (boolean). That covers
/// .xlam workbook add-ins and the built-in Excel add-ins (Solver, etc.).
/// </para>
/// <para>
/// What's NOT available on Mac that the Windows version exposes:
/// </para>
/// <list type="bullet">
///   <item><c>COMAddIns</c> — no COM model on macOS; VSTO/COM add-ins
///         don't exist in the Mac Office runtime.</item>
///   <item>XLL native add-ins — Mac Excel doesn't support XLLs at all
///         (Microsoft removed the loader from Mac Office 2016+).</item>
/// </list>
/// <para>
/// So this command's output on Mac is strictly a subset of the Windows
/// equivalent — only the workbook/Native add-ins surface, with the
/// <c>kind</c> column always reading <c>VBA</c> (.xlam) or <c>NATIVE</c>.
/// Same TSV column shape as the Windows version so a single ScripTree
/// app calling <c>combridge excel list-addins</c> works on both OSes.
/// </para>
/// </remarks>
internal sealed class ExcelMacListAddinsCommand : IBridgeCommand
{
    public string Name => "list-addins";
    public string Usage => "list-addins   (lists Mac Excel add-ins via AppleScript; subset of Windows)";

    public Task<int> RunAsync(object comRoot, string[] args, TextWriter output)
    {
        output.WriteLine("# columns: name\tid\tloaded\tkind\tdescription");
        if (!Osascript.IsAvailable())
        {
            output.WriteLine("# WARN: osascript not available — cannot enumerate add-ins.");
            output.WriteLine();
            output.WriteLine("# total: 0");
            return Task.FromResult(0);
        }

        // Single AppleScript invocation returns one line per add-in,
        // FIELD_SEP-delimited so we can split safely on the C# side
        // even if names contain spaces. Using U+241E (RS) for the same
        // reasoning as the Outlook search command.
        const string FieldSep = "␞";
        string script = $@"
tell application ""Microsoft Excel""
    set out to {{}}
    try
        repeat with a in addins
            try
                set rowText to (name of a as text) & ""{FieldSep}"" & (full name of a as text) & ""{FieldSep}"" & (installed of a as text)
                set end of out to rowText
            end try
        end repeat
    end try
    set AppleScript's text item delimiters to linefeed
    return out as text
end tell".Trim();

        string? raw;
        try { raw = Osascript.TryRun(script); }
        catch { raw = null; }
        if (string.IsNullOrEmpty(raw))
        {
            output.WriteLine();
            output.WriteLine("# total: 0");
            return Task.FromResult(0);
        }

        int count = 0;
        foreach (var line in raw.Split('\n', StringSplitOptions.RemoveEmptyEntries))
        {
            var parts = line.Split(FieldSep);
            if (parts.Length < 3) continue;
            string name      = EscTab(parts[0]);
            string fullName  = EscTab(parts[1]);
            string installed = parts[2].Trim().Equals("true", StringComparison.OrdinalIgnoreCase) ? "true" : "false";
            string kind      = ClassifyKind(fullName);
            output.WriteLine($"{name}\t{fullName}\t{installed}\t{kind}\t");
            count++;
        }

        output.WriteLine();
        output.WriteLine($"# total: {count} (Mac AppleScript surface — no COM/XLL on macOS)");
        return Task.FromResult(0);
    }

    private static string EscTab(string? s)
        => (s ?? "").Replace('\t', ' ').Replace('\n', ' ').Replace('\r', ' ');

    private static string ClassifyKind(string? fullPath)
    {
        var ext = Path.GetExtension(fullPath ?? "").ToLowerInvariant();
        return ext switch
        {
            ".xla" or ".xlam" => "VBA",
            "" => "NATIVE",
            _ => ext.TrimStart('.').ToUpperInvariant()
        };
    }
}
