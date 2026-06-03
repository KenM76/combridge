using ComBridge.Core;
using Ol = global::Microsoft.Office.Interop.Outlook;

namespace ComBridge.Plugins.Outlook;

/// <summary>
/// <c>outlook list-addins</c> — enumerate Outlook COM add-ins as TSV.
/// </summary>
/// <remarks>
/// Outlook's add-in model is COM-only (no equivalent to Word's classic
/// templates or Excel's XLLs). Older builds expose <c>COMAddIns</c>
/// freely; newer security-hardened deployments (Office 365 modern
/// hardening, GPO-locked tenants) may restrict access — we tolerate
/// with a WARN row rather than failing the command, same convention
/// as the other plugins' <c>list-addins</c>.
/// </remarks>
internal sealed class OlListAddinsCommand : IBridgeCommand
{
    public string Name => "list-addins";
    public string Usage => "list-addins   (lists Outlook COM add-ins as TSV)";

    public Task<int> RunAsync(object comRoot, string[] args, TextWriter output)
    {
        var app = (Ol._Application)comRoot;
        output.WriteLine("# columns: name\tid\tloaded\tkind\tdescription");

        int count = 0;
        try
        {
            foreach (Microsoft.Office.Core.COMAddIn ca in app.COMAddIns)
            {
                output.WriteLine(string.Join("\t",
                    OlAddinFormat.EscTab(ca.Description),
                    OlAddinFormat.EscTab(ca.ProgId),
                    ca.Connect.ToString().ToLowerInvariant(),
                    "COM",
                    OlAddinFormat.EscTab(OlAddinFormat.SafeProperty(() => ca.Guid))));
                count++;
            }
        }
        catch (Exception ex) { output.WriteLine($"# WARN: COMAddIns enumeration failed: {ex.Message}"); }

        output.WriteLine();
        output.WriteLine($"# total: {count}");
        return Task.FromResult(0);
    }
}

internal static class OlAddinFormat
{
    public static string EscTab(string? s)
        => (s ?? "").Replace('\t', ' ').Replace('\n', ' ').Replace('\r', ' ');
    public static string SafeProperty(Func<string?> read)
    { try { return read() ?? ""; } catch { return ""; } }
}
