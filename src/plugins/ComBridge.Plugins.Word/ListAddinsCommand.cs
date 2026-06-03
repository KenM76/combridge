using ComBridge.Core;
using Wd = global::Microsoft.Office.Interop.Word;

namespace ComBridge.Plugins.Word;

/// <summary>
/// <c>word list-addins</c> — enumerate Word's COM add-ins and classic
/// add-ins (global templates / WLLs) in one TSV stream. Same diagnostic
/// role as <c>list-sessions</c> / <c>info</c>.
/// </summary>
/// <remarks>
/// Word splits add-ins across two collections:
/// <list type="bullet">
///   <item><c>Application.COMAddIns</c> — VSTO / COM add-ins (Mendeley,
///         Grammarly, etc.). Per item: <c>ProgId</c>, <c>Description</c>,
///         <c>Connect</c>.</item>
///   <item><c>Application.AddIns</c> — global templates (.dot/.dotm) and
///         Word add-ins (.wll). Per item: <c>Name</c>, <c>Path</c>,
///         <c>Installed</c>, <c>Autoload</c>.</item>
/// </list>
/// Each collection is wrapped in its own try/catch — security policies
/// occasionally deny access to one but not the other.
/// </remarks>
internal sealed class WdListAddinsCommand : IBridgeCommand
{
    public string Name => "list-addins";
    public string Usage => "list-addins   (lists COM + classic add-ins as TSV)";

    public Task<int> RunAsync(object comRoot, string[] args, TextWriter output)
    {
        var app = (Wd._Application)comRoot;
        output.WriteLine("# columns: name\tid\tloaded\tkind\tdescription");

        int count = 0;
        try
        {
            foreach (Microsoft.Office.Core.COMAddIn ca in app.COMAddIns)
            {
                output.WriteLine(string.Join("\t",
                    WdAddinFormat.EscTab(ca.Description),
                    WdAddinFormat.EscTab(ca.ProgId),
                    ca.Connect.ToString().ToLowerInvariant(),
                    "COM",
                    WdAddinFormat.EscTab(WdAddinFormat.SafeProperty(() => ca.Guid))));
                count++;
            }
        }
        catch (Exception ex) { output.WriteLine($"# WARN: COMAddIns enumeration failed: {ex.Message}"); }

        try
        {
            foreach (Wd.AddIn ai in app.AddIns)
            {
                string fullPath = WdAddinFormat.SafeProperty(() => ai.Path)
                                + Path.DirectorySeparatorChar
                                + ai.Name;
                output.WriteLine(string.Join("\t",
                    WdAddinFormat.EscTab(ai.Name),
                    WdAddinFormat.EscTab(fullPath),
                    ai.Installed.ToString().ToLowerInvariant(),
                    ClassifyKind(ai.Name),
                    WdAddinFormat.EscTab(WdAddinFormat.SafeProperty(() => ai.Autoload ? "autoload" : ""))));
                count++;
            }
        }
        catch (Exception ex) { output.WriteLine($"# WARN: AddIns enumeration failed: {ex.Message}"); }

        output.WriteLine();
        output.WriteLine($"# total: {count}");
        return Task.FromResult(0);
    }

    private static string ClassifyKind(string? name)
    {
        var ext = Path.GetExtension(name ?? "").ToLowerInvariant();
        return ext switch
        {
            ".wll" => "WLL",
            ".dot" or ".dotm" or ".dotx" => "TEMPLATE",
            ".dll" => "COM",
            _ => ext.TrimStart('.').ToUpperInvariant()
        };
    }
}

internal static class WdAddinFormat
{
    public static string EscTab(string? s)
        => (s ?? "").Replace('\t', ' ').Replace('\n', ' ').Replace('\r', ' ');
    public static string SafeProperty(Func<string?> read)
    { try { return read() ?? ""; } catch { return ""; } }
}
