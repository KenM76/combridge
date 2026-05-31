using System.Globalization;
using System.Runtime.InteropServices;
using ComBridge.Core;
using ComBridge.Mac.Common;
using Microsoft.CodeAnalysis;

namespace ComBridge.Plugins.PowerPoint.Mac;

public sealed class PowerPointMacPlugin : IComBridgePlugin
{
    public string Name => "powerpoint";
    public string Description => "Microsoft PowerPoint for macOS (AppleScript backend). Globals: pptApp.";
    public string[] ProgIds => new[] { "Microsoft PowerPoint" };
    public bool AllowCreateNew => true;
    public Type GlobalsType => typeof(PptMacGlobals);

    public IReadOnlyCollection<OSPlatform> SupportedPlatforms => new[] { OSPlatform.OSX };

    public object CreateGlobals(object comRoot) => new PptMacGlobals();

    public IEnumerable<MetadataReference> ScriptReferences
    {
        get
        {
            yield return MetadataReference.CreateFromFile(typeof(PowerPointMacPlugin).Assembly.Location);
            yield return MetadataReference.CreateFromFile(typeof(Osascript).Assembly.Location);
        }
    }

    public IEnumerable<string> ScriptUsings => new[] { "ComBridge.Plugins.PowerPoint.Mac" };

    public IEnumerable<IBridgeCommand> Commands => new IBridgeCommand[]
    {
        new PptMacInfoCommand(),
        new PptMacListSlidesCommand(),
    };

    public List<(object Root, SessionInfo Info)> FindSessions()
    {
        var sessions = new List<(object, SessionInfo)>();
        if (!Osascript.IsAvailable()) return sessions;

        var running = Osascript.TryRun(
            "tell application \"System Events\" to (name of processes) contains \"Microsoft PowerPoint\"");
        if (running != "true") return sessions;

        int? pid = null;
        var pidRaw = Osascript.TryRun(
            "tell application \"System Events\" to unix id of (first process whose name is \"Microsoft PowerPoint\")");
        if (int.TryParse(pidRaw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n)) pid = n;

        string? title = Osascript.TryRun("tell application \"Microsoft PowerPoint\" to name of active presentation");
        if (string.IsNullOrEmpty(title)) title = null;

        var desc = (pid, title) switch
        {
            (int pp, string t) when !string.IsNullOrEmpty(t) => $"pid={pp}  title={t}",
            (int pp, _)                                       => $"pid={pp}",
            (null, string t) when !string.IsNullOrEmpty(t)    => t,
            _                                                  => "(no info)",
        };

        sessions.Add((new object(), new SessionInfo(1, pid, title, desc)));
        return sessions;
    }

    public (int? Pid, string? Title) DescribeInstance(object comRoot) => (null, null);
}

internal sealed class PptMacInfoCommand : IBridgeCommand
{
    public string Name => "info";
    public string Usage => "info   (prints PowerPoint version + active presentation)";

    public Task<int> RunAsync(object comRoot, string[] args, TextWriter output)
    {
        var app = new PptMacApp();
        try { output.WriteLine($"PowerPoint version: {app.Version}"); }
        catch { output.WriteLine("PowerPoint version: (not running)"); }
        try { output.WriteLine($"Visible:            {app.Visible}"); } catch { }
        try { output.WriteLine($"ActivePresentation: {app.ActivePresentationName ?? "(none)"}"); } catch { }
        try { output.WriteLine($"Path:               {app.ActivePresentationPath ?? "(unsaved)"}"); } catch { }
        try { output.WriteLine($"Presentations:      {app.PresentationCount}"); } catch { }
        try { output.WriteLine($"Slides (active):    {app.SlideCount}"); } catch { }
        try { output.WriteLine($"ActiveSlide:        {app.ActiveSlideIndex?.ToString() ?? "(none)"}"); } catch { }
        return Task.FromResult(0);
    }
}

internal sealed class PptMacListSlidesCommand : IBridgeCommand
{
    public string Name => "list-slides";
    public string Usage => "list-slides   (prints index + title of every slide in the active presentation)";

    public Task<int> RunAsync(object comRoot, string[] args, TextWriter output)
    {
        var app = new PptMacApp();
        try
        {
            var titles = app.SlideTitles;
            for (int i = 0; i < titles.Length; i++)
            {
                output.WriteLine($"  [{i + 1,3}] {(string.IsNullOrEmpty(titles[i]) ? "(no title)" : titles[i])}");
            }
            output.WriteLine($"\n{titles.Length} slide(s).");
            return Task.FromResult(0);
        }
        catch (OsascriptException ex)
        {
            output.WriteLine($"ERROR: {ex.Message}");
            return Task.FromResult(4);
        }
    }
}
