using ComBridge.Core;
using Microsoft.CodeAnalysis;
using Pp = global::Microsoft.Office.Interop.PowerPoint;

namespace ComBridge.Plugins.PowerPoint;

/// <summary>
/// Globals exposed to user .csx scripts. <c>pptApp</c> is the
/// <c>_Application</c> dispinterface to match what ROT-fetched COM objects
/// expose (same reasoning as the Excel/Word plugins).
/// </summary>
public sealed class PptGlobals
{
    public Pp._Application pptApp { get; }
    public Pp.Presentation? pptPres { get; }
    public Pp.Slide? pptSlide { get; }

    internal PptGlobals(Pp._Application app)
    {
        pptApp = app;
        try { pptPres = app.ActivePresentation; } catch { pptPres = null; }
        try
        {
            // ActiveWindow can be null (no presentation open). View.Slide returns
            // 'object' because in SlideShow mode it's a different Slide type — cast.
            var win = app.ActiveWindow;
            pptSlide = win?.View?.Slide as Pp.Slide;
        }
        catch { pptSlide = null; }
    }
}

public sealed class PowerPointPlugin : IComBridgePlugin
{
    public string Name => "powerpoint";
    public string Description => "Microsoft PowerPoint (attach via ROT or launch new). Globals: pptApp, pptPres, pptSlide.";
    public string[] ProgIds => new[] { "PowerPoint.Application" };
    public bool AllowCreateNew => true;
    public Type GlobalsType => typeof(PptGlobals);

    public object CreateGlobals(object comRoot) => new PptGlobals((Pp._Application)comRoot);

    public IEnumerable<MetadataReference> ScriptReferences
    {
        get
        {
            var here = Path.GetDirectoryName(typeof(PowerPointPlugin).Assembly.Location)!;
            foreach (var dll in Directory.EnumerateFiles(here, "Microsoft.Office.Interop.PowerPoint*.dll"))
                yield return MetadataReference.CreateFromFile(dll);
            foreach (var dll in Directory.EnumerateFiles(here, "office.dll"))
                yield return MetadataReference.CreateFromFile(dll);
        }
    }

    public IEnumerable<string> ScriptUsings => new[] { "Microsoft.Office.Interop.PowerPoint" };

    public IEnumerable<IBridgeCommand> Commands => new IBridgeCommand[]
    {
        new PptInfoCommand(),
    };

    // PowerPoint per-presentation file monikers in the ROT.
    public IEnumerable<string> RotMonikerPatterns => new[]
    {
        @"\.(pptx|ppt|pptm|potx|potm|ppsx|ppsm|pps|ppam)$",
    };

    public object? TryExtractRoot(object monikerBound)
    {
        try
        {
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
    /// PowerPoint exposes <c>_Application.HWND</c> directly (note uppercase HWND).
    /// </summary>
    public (int? Pid, string? Title) DescribeInstance(object comRoot)
    {
        try
        {
            var app = (Pp._Application)comRoot;
            int? pid = null;
            try
            {
                var hwnd = (IntPtr)app.HWND;
                pid = SessionPicker.PidFromHwnd(hwnd);
            }
            catch { }

            string? title = null;
            try { title = app.ActivePresentation?.Name; } catch { }
            return (pid, title);
        }
        catch
        {
            return (null, null);
        }
    }
}

internal sealed class PptInfoCommand : IBridgeCommand
{
    public string Name => "info";
    public string Usage => "info   (prints PowerPoint version + active presentation)";

    public Task<int> RunAsync(object comRoot, string[] args, TextWriter output)
    {
        var app = (Pp._Application)comRoot;
        output.WriteLine($"PowerPoint version: {app.Version}");
        try { output.WriteLine($"Visible:            {app.Visible}"); } catch { output.WriteLine("Visible:            (unavailable)"); }
        try { output.WriteLine($"ActivePresentation: {app.ActivePresentation?.Name ?? "(none)"}"); }
        catch { output.WriteLine("ActivePresentation: (none)"); }
        try { output.WriteLine($"Presentations:      {app.Presentations.Count}"); } catch { }
        try
        {
            var win = app.ActiveWindow;
            var slide = win?.View?.Slide as Pp.Slide;
            output.WriteLine($"ActiveSlide:        {slide?.SlideIndex ?? -1}");
        }
        catch { output.WriteLine("ActiveSlide:        (unavailable)"); }
        return Task.FromResult(0);
    }
}
