using ComBridge.Core;
using Microsoft.CodeAnalysis;
using Wd = global::Microsoft.Office.Interop.Word;

namespace ComBridge.Plugins.Word;

/// <summary>
/// Globals exposed to user .csx scripts. Cast to the <c>_Application</c>
/// dispinterface (NOT the <c>Application</c> co-class) so ROT-fetched COM
/// objects work correctly — see the Excel plugin's notes for the rationale.
/// </summary>
public sealed class WdGlobals
{
    public Wd._Application wdApp { get; }
    public Wd.Document? wdDoc { get; }

    internal WdGlobals(Wd._Application app)
    {
        wdApp = app;
        try { wdDoc = app.ActiveDocument; } catch { wdDoc = null; }
    }
}

public sealed class WordPlugin : IComBridgePlugin
{
    public string Name => "word";
    public string Description => "Microsoft Word (attach via ROT or launch new). Globals: wdApp, wdDoc.";
    public string[] ProgIds => new[] { "Word.Application" };
    public bool AllowCreateNew => true;
    public Type GlobalsType => typeof(WdGlobals);

    public object CreateGlobals(object comRoot) => new WdGlobals((Wd._Application)comRoot);

    public IEnumerable<MetadataReference> ScriptReferences
    {
        get
        {
            var here = Path.GetDirectoryName(typeof(WordPlugin).Assembly.Location)!;
            foreach (var dll in Directory.EnumerateFiles(here, "Microsoft.Office.Interop.Word*.dll"))
                yield return MetadataReference.CreateFromFile(dll);
            foreach (var dll in Directory.EnumerateFiles(here, "office.dll"))
                yield return MetadataReference.CreateFromFile(dll);
        }
    }

    public IEnumerable<string> ScriptUsings => new[] { "Microsoft.Office.Interop.Word" };

    public IEnumerable<IBridgeCommand> Commands => new IBridgeCommand[]
    {
        new WdInfoCommand(),
    };

    // Word registers per-document file monikers in the ROT, same as Excel.
    // We match document/template/RTF extensions and ascend Document→Application
    // in TryExtractRoot. PID dedupe in SessionPicker collapses multiple docs
    // in one WINWORD.EXE to a single session.
    public IEnumerable<string> RotMonikerPatterns => new[]
    {
        @"\.(docx|doc|docm|dotx|dotm|rtf)$",
    };

    public object? TryExtractRoot(object monikerBound)
    {
        try
        {
            // monikerBound is a Word Document RCW; ascend to its parent Application.
            dynamic d = monikerBound;
            object? app = d.Application;
            return app;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Word's main-window HWND lives on <c>_Application.ActiveWindow.Hwnd</c>.
    /// If no document is open, ActiveWindow is null — we fall back to the
    /// first entry in the Windows collection. If both fail, return (null, title)
    /// so SessionPicker still keeps the entry (title alone is enough).
    /// </summary>
    public (int? Pid, string? Title) DescribeInstance(object comRoot)
    {
        try
        {
            var app = (Wd._Application)comRoot;
            int? pid = null;
            try
            {
                Wd.Window? win = null;
                try { win = app.ActiveWindow; } catch { }
                if (win is null)
                {
                    try { if (app.Windows.Count > 0) win = app.Windows[1]; } catch { }
                }
                if (win is not null)
                {
                    var hwnd = (IntPtr)win.Hwnd;
                    pid = SessionPicker.PidFromHwnd(hwnd);
                }
            }
            catch { /* tolerate */ }

            string? title = null;
            try { title = app.ActiveDocument?.Name; } catch { }
            return (pid, title);
        }
        catch
        {
            return (null, null);
        }
    }
}

internal sealed class WdInfoCommand : IBridgeCommand
{
    public string Name => "info";
    public string Usage => "info   (prints Word version + active document/window)";

    public Task<int> RunAsync(object comRoot, string[] args, TextWriter output)
    {
        var app = (Wd._Application)comRoot;
        output.WriteLine($"Word version:  {app.Version}");
        try { output.WriteLine($"Visible:       {app.Visible}"); } catch { output.WriteLine("Visible:       (unavailable)"); }
        try { output.WriteLine($"ActiveDoc:     {app.ActiveDocument?.Name ?? "(none)"}"); }
        catch { output.WriteLine("ActiveDoc:     (no document open)"); }
        try { output.WriteLine($"Documents:     {app.Documents.Count}"); } catch { }
        return Task.FromResult(0);
    }
}
