using ComBridge.Core;
using Pp = global::Microsoft.Office.Interop.PowerPoint;

namespace ComBridge.Plugins.PowerPoint;

/// <summary>
/// <c>powerpoint list-addins</c> — enumerate PowerPoint COM + classic
/// add-ins (.ppam / .ppa) as TSV. Diagnostic infrastructure, same shape
/// as the other plugins' <c>list-addins</c>.
/// </summary>
/// <remarks>
/// PowerPoint exposes <c>AddIns</c> entries with <c>Loaded</c> (MsoTriState)
/// rather than a bool, because PowerPoint distinguishes "registered but
/// not loaded this session" from "registered and currently loaded."
/// </remarks>
internal sealed class PptListAddinsCommand : IBridgeCommand
{
    public string Name => "list-addins";
    public string Usage => "list-addins   (lists COM + .ppam/.ppa add-ins as TSV)";

    public Task<int> RunAsync(object comRoot, string[] args, TextWriter output)
    {
        var app = (Pp._Application)comRoot;
        output.WriteLine("# columns: name\tid\tloaded\tkind\tdescription");

        int count = 0;
        try
        {
            foreach (Microsoft.Office.Core.COMAddIn ca in app.COMAddIns)
            {
                output.WriteLine(string.Join("\t",
                    PptAddinFormat.EscTab(ca.Description),
                    PptAddinFormat.EscTab(ca.ProgId),
                    ca.Connect.ToString().ToLowerInvariant(),
                    "COM",
                    PptAddinFormat.EscTab(PptAddinFormat.SafeProperty(() => ca.Guid))));
                count++;
            }
        }
        catch (Exception ex) { output.WriteLine($"# WARN: COMAddIns enumeration failed: {ex.Message}"); }

        try
        {
            foreach (Pp.AddIn ai in app.AddIns)
            {
                output.WriteLine(string.Join("\t",
                    PptAddinFormat.EscTab(ai.Name),
                    PptAddinFormat.EscTab(PptAddinFormat.SafeProperty(() => ai.FullName)),
                    PptAddinFormat.SafeProperty(() => (ai.Loaded == Microsoft.Office.Core.MsoTriState.msoTrue) ? "true" : "false"),
                    ClassifyKind(PptAddinFormat.SafeProperty(() => ai.FullName)),
                    PptAddinFormat.EscTab(PptAddinFormat.SafeProperty(() => (ai.Registered == Microsoft.Office.Core.MsoTriState.msoTrue) ? "registered" : ""))));
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
            ".ppam" or ".ppa" => "VBA",
            ".dll" => "COM",
            _ => ext.TrimStart('.').ToUpperInvariant()
        };
    }
}

internal static class PptAddinFormat
{
    public static string EscTab(string? s)
        => (s ?? "").Replace('\t', ' ').Replace('\n', ' ').Replace('\r', ' ');
    public static string SafeProperty(Func<string?> read)
    { try { return read() ?? ""; } catch { return ""; } }
}
