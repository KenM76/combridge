using ComBridge.Core;
using Xl = global::Microsoft.Office.Interop.Excel;

namespace ComBridge.Plugins.Excel;

/// <summary>
/// <c>excel list-addins</c> — enumerate Excel COM add-ins AND classic
/// add-ins (XLL, .xla/.xlam, automation) in one TSV stream. Same diagnostic
/// role as <c>list-sessions</c> / <c>info</c> — universal infrastructure,
/// not domain logic.
/// </summary>
/// <remarks>
/// <para>
/// Excel splits add-ins across two distinct collections that consumers
/// frequently confuse:
/// </para>
/// <list type="bullet">
///   <item><c>Application.COMAddIns</c> — registered COM/VSTO add-ins.
///         Per item: <c>ProgId</c>, <c>Description</c>, <c>Guid</c>,
///         <c>Connect</c> (whether currently loaded into the running
///         instance). A "Connect=false" entry is installed but not
///         currently loaded.</item>
///   <item><c>Application.AddIns</c> — classic .xla/.xlam workbook
///         add-ins, .xll automation add-ins, and built-in Office add-ins
///         (Solver, Analysis ToolPak, etc.). Per item: <c>Name</c>,
///         <c>FullName</c> (path), <c>Title</c>, <c>Installed</c>.</item>
/// </list>
/// <para>
/// We emit BOTH in one TSV with a <c>kind</c> column distinguishing
/// COM / XLL / VBA (.xla/.xlam) / NATIVE so the consumer can filter.
/// Each collection is wrapped in its own try/catch so a partial
/// enumeration failure (security policy denying access to one collection)
/// still emits the other.
/// </para>
/// </remarks>
internal sealed class ListAddinsCommand : IBridgeCommand
{
    public string Name => "list-addins";
    public string Usage => "list-addins   (lists COM + XLL + .xla/.xlam add-ins as TSV)";

    public Task<int> RunAsync(object comRoot, string[] args, TextWriter output)
    {
        var app = (Xl._Application)comRoot;
        output.WriteLine("# columns: name\tid\tloaded\tkind\tdescription");

        int count = 0;
        try
        {
            foreach (Microsoft.Office.Core.COMAddIn ca in app.COMAddIns)
            {
                output.WriteLine(string.Join("\t",
                    AddinFormat.EscTab(ca.Description),
                    AddinFormat.EscTab(ca.ProgId),
                    ca.Connect.ToString().ToLowerInvariant(),
                    "COM",
                    AddinFormat.EscTab(AddinFormat.SafeProperty(() => ca.Guid))));
                count++;
            }
        }
        catch (Exception ex) { output.WriteLine($"# WARN: COMAddIns enumeration failed: {ex.Message}"); }

        try
        {
            foreach (Xl.AddIn ai in app.AddIns)
            {
                output.WriteLine(string.Join("\t",
                    AddinFormat.EscTab(ai.Name),
                    AddinFormat.EscTab(AddinFormat.SafeProperty(() => ai.FullName)),
                    ai.Installed.ToString().ToLowerInvariant(),
                    ClassifyKind(AddinFormat.SafeProperty(() => ai.FullName)),
                    AddinFormat.EscTab(AddinFormat.SafeProperty(() => ai.Title))));
                count++;
            }
        }
        catch (Exception ex) { output.WriteLine($"# WARN: AddIns enumeration failed: {ex.Message}"); }

        output.WriteLine();
        output.WriteLine($"# total: {count}");
        return Task.FromResult(0);
    }

    private static string ClassifyKind(string? fullPath)
    {
        var ext = Path.GetExtension(fullPath ?? "").ToLowerInvariant();
        return ext switch
        {
            ".xll" => "XLL",
            ".xla" or ".xlam" => "VBA",
            ".dll" => "COM",
            "" => "NATIVE",
            _ => ext.TrimStart('.').ToUpperInvariant()
        };
    }
}

/// <summary>
/// TSV-row formatting helpers shared by every plugin's <c>list-addins</c>
/// implementation. Kept here (not in Core) because they're a hair too
/// minimal to deserve their own assembly — same pattern as the
/// <c>ScriptScaffold</c> exception, but going the other direction: shared
/// inside one plugin assembly via internal static class.
/// </summary>
internal static class AddinFormat
{
    /// <summary>
    /// Replace tabs/newlines with spaces so a TSV row stays one logical line.
    /// We do NOT JSON-escape — TSV is the contract; consumers split on \t.
    /// </summary>
    public static string EscTab(string? s)
        => (s ?? "").Replace('\t', ' ').Replace('\n', ' ').Replace('\r', ' ');

    /// <summary>
    /// Tolerantly read a property that may throw on certain add-ins
    /// (e.g. some COM add-ins refuse to expose <c>Guid</c>; some XLLs
    /// have null <c>FullName</c>).
    /// </summary>
    public static string SafeProperty(Func<string?> read)
    { try { return read() ?? ""; } catch { return ""; } }
}
