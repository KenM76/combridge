using ComBridge.Core;
using ComBridge.Mac.Common;

namespace ComBridge.Plugins.Word.Mac;

/// <summary>
/// <c>word list-addins</c> for macOS — best-effort enumeration via
/// AppleScript's <c>add-ins</c> collection on Microsoft Word.
/// </summary>
/// <remarks>
/// <para>
/// Mac Word's AppleScript dictionary exposes <c>add-ins of application</c>
/// returning a list of <c>add-in</c> records, each with <c>name</c>,
/// <c>path</c>, and <c>installed</c> (boolean). That covers global
/// templates (.dotm / .dotx) and any Word add-ins the user has
/// registered manually. <c>WLL</c> add-ins don't exist on Mac (Word
/// add-in DLL hosting was Windows-only).
/// </para>
/// <para>
/// What's NOT available on Mac that the Windows version exposes:
/// COM/VSTO add-ins (Mendeley, Grammarly, etc.) — no COM model on macOS.
/// Output is a strict subset of the Windows equivalent, but uses the
/// same TSV column shape so a single ScripTree app works cross-OS.
/// </para>
/// </remarks>
internal sealed class WordMacListAddinsCommand : IBridgeCommand
{
    public string Name => "list-addins";
    public string Usage => "list-addins   (lists Mac Word add-ins via AppleScript; subset of Windows)";

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

        const string FieldSep = "␞";
        string script = $@"
tell application ""Microsoft Word""
    set out to {{}}
    try
        repeat with a in add-ins
            try
                set rowText to (name of a as text) & ""{FieldSep}"" & (path of a as text) & ""{FieldSep}"" & (installed of a as text)
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
            string fullPath  = EscTab(parts[1]) + Path.DirectorySeparatorChar + parts[0];
            string installed = parts[2].Trim().Equals("true", StringComparison.OrdinalIgnoreCase) ? "true" : "false";
            string kind      = ClassifyKind(name);
            output.WriteLine($"{name}\t{fullPath}\t{installed}\t{kind}\t");
            count++;
        }

        output.WriteLine();
        output.WriteLine($"# total: {count} (Mac AppleScript surface — no COM/WLL on macOS)");
        return Task.FromResult(0);
    }

    private static string EscTab(string? s)
        => (s ?? "").Replace('\t', ' ').Replace('\n', ' ').Replace('\r', ' ');

    private static string ClassifyKind(string? name)
    {
        var ext = Path.GetExtension(name ?? "").ToLowerInvariant();
        return ext switch
        {
            ".dot" or ".dotm" or ".dotx" => "TEMPLATE",
            _ => ext.TrimStart('.').ToUpperInvariant()
        };
    }
}
