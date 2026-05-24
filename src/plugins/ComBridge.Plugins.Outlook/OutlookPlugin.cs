using ComBridge.Core;
using Microsoft.CodeAnalysis;
using Ol = global::Microsoft.Office.Interop.Outlook;

namespace ComBridge.Plugins.Outlook;

/// <summary>
/// Globals exposed to user .csx scripts.
/// <para>
/// Outlook is fundamentally different from the document-based Office apps:
/// one MAPI session per user, no "open documents" concept, scripts almost
/// always go through the <c>NameSpace</c> (typically <c>"MAPI"</c>) to
/// access folders, items, and accounts.
/// </para>
/// </summary>
public sealed class OlGlobals
{
    public Ol._Application olApp { get; }
    public Ol.NameSpace olNs { get; }
    public Ol.Explorer? olExplorer { get; }

    internal OlGlobals(Ol._Application app)
    {
        olApp = app;
        olNs = app.GetNamespace("MAPI");
        try { olExplorer = app.ActiveExplorer(); } catch { olExplorer = null; }
    }
}

public sealed class OutlookPlugin : IComBridgePlugin
{
    public string Name => "outlook";
    public string Description => "Microsoft Outlook (single MAPI session). Globals: olApp, olNs, olExplorer.";
    public string[] ProgIds => new[] { "Outlook.Application" };
    public bool AllowCreateNew => true;
    public Type GlobalsType => typeof(OlGlobals);

    public object CreateGlobals(object comRoot) => new OlGlobals((Ol._Application)comRoot);

    public IEnumerable<MetadataReference> ScriptReferences
    {
        get
        {
            var here = Path.GetDirectoryName(typeof(OutlookPlugin).Assembly.Location)!;
            foreach (var dll in Directory.EnumerateFiles(here, "Microsoft.Office.Interop.Outlook*.dll"))
                yield return MetadataReference.CreateFromFile(dll);
            foreach (var dll in Directory.EnumerateFiles(here, "office.dll"))
                yield return MetadataReference.CreateFromFile(dll);
        }
    }

    public IEnumerable<string> ScriptUsings => new[] { "Microsoft.Office.Interop.Outlook" };

    public IEnumerable<IBridgeCommand> Commands => new IBridgeCommand[]
    {
        new OlInfoCommand(),
    };

    // Outlook has no per-document moniker concept (it's a MAPI session, not
    // a document app). The only ROT entry is the class moniker, which Outlook
    // DOES register reliably — so an empty RotMonikerPatterns + the
    // GetActiveObject fallback in SessionPicker is sufficient.
    public IEnumerable<string> RotMonikerPatterns => Array.Empty<string>();

    // Outlook's main window HWND isn't directly on Application. Use the
    // ActiveExplorer's CommandBars host or just resolve PID via process name —
    // single MAPI session means there's only one OUTLOOK.EXE anyway.
    public (int? Pid, string? Title) DescribeInstance(object comRoot)
    {
        try
        {
            var app = (Ol._Application)comRoot;
            int? pid = null;
            try
            {
                // Outlook 2010+ exposes Explorer.Caption but not Hwnd directly.
                // Most reliable PID source: enumerate OUTLOOK.EXE processes
                // (there's only ever one per user session).
                var procs = System.Diagnostics.Process.GetProcessesByName("OUTLOOK");
                if (procs.Length > 0) pid = procs[0].Id;
            }
            catch { }

            string? title = null;
            try
            {
                var exp = app.ActiveExplorer();
                if (exp is not null)
                {
                    var folder = exp.CurrentFolder;
                    title = folder?.Name ?? exp.Caption;
                }
                title ??= $"Outlook v{app.Version}";
            }
            catch
            {
                try { title = $"Outlook v{app.Version}"; } catch { }
            }
            return (pid, title);
        }
        catch
        {
            return (null, null);
        }
    }
}

internal sealed class OlInfoCommand : IBridgeCommand
{
    public string Name => "info";
    public string Usage => "info   (prints Outlook version + default folders + active explorer)";

    public Task<int> RunAsync(object comRoot, string[] args, TextWriter output)
    {
        var app = (Ol._Application)comRoot;
        output.WriteLine($"Outlook version: {app.Version}");
        try
        {
            var ns = app.GetNamespace("MAPI");
            output.WriteLine($"User:            {ns.CurrentUser?.Name ?? "(unknown)"}");

            // Default Inbox item count is a useful quick sanity check.
            var inbox = ns.GetDefaultFolder(Ol.OlDefaultFolders.olFolderInbox);
            output.WriteLine($"Inbox items:     {inbox.Items.Count}");

            // List top-level stores (accounts).
            output.WriteLine("Stores:");
            for (int i = 1; i <= ns.Stores.Count; i++)
            {
                var store = ns.Stores[i];
                output.WriteLine($"  - {store.DisplayName}");
            }
        }
        catch (Exception ex)
        {
            output.WriteLine($"(namespace inspection failed: {ex.Message})");
        }

        try
        {
            var exp = app.ActiveExplorer();
            if (exp is not null)
            {
                output.WriteLine($"ActiveExplorer:  {exp.Caption}");
                output.WriteLine($"CurrentFolder:   {exp.CurrentFolder?.Name ?? "(none)"}");
            }
            else
            {
                output.WriteLine("ActiveExplorer:  (none — window may be minimized)");
            }
        }
        catch { output.WriteLine("ActiveExplorer:  (unavailable)"); }

        return Task.FromResult(0);
    }
}
