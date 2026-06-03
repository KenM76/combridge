using System.Globalization;
using System.Runtime.InteropServices;
using ComBridge.Core;
using ComBridge.Mac.Common;
using Microsoft.CodeAnalysis;

namespace ComBridge.Plugins.Word.Mac;

/// <summary>
/// macOS Word plugin. Same CLI contract as the Windows Word plugin
/// (<c>Name = "word"</c>; <c>info</c>, <c>extract-text</c>, <c>doc-stats</c>
/// commands), implemented via AppleScript (<c>osascript</c>) instead of
/// COM. PluginLoader's SupportedPlatforms filter ensures only one "word"
/// plugin loads per OS.
/// </summary>
public sealed class WordMacPlugin : IComBridgePlugin
{
    public string Name => "word";
    public string Description => "Microsoft Word for macOS (AppleScript backend). Globals: wdApp.";
    public string[] ProgIds => new[] { "Microsoft Word" };
    public bool AllowCreateNew => true;
    public Type GlobalsType => typeof(WdMacGlobals);

    public IReadOnlyCollection<OSPlatform> SupportedPlatforms => new[] { OSPlatform.OSX };

    public object CreateGlobals(object comRoot) => new WdMacGlobals();

    public IEnumerable<MetadataReference> ScriptReferences
    {
        get
        {
            yield return MetadataReference.CreateFromFile(typeof(WordMacPlugin).Assembly.Location);
            yield return MetadataReference.CreateFromFile(typeof(Osascript).Assembly.Location);
        }
    }

    public IEnumerable<string> ScriptUsings => new[] { "ComBridge.Plugins.Word.Mac" };

    public IEnumerable<IBridgeCommand> Commands => new IBridgeCommand[]
    {
        new WordMacInfoCommand(),
        new WordMacExtractTextCommand(),
        new WordMacDocStatsCommand(),
        new WordMacListAddinsCommand(),
    };

    public List<(object Root, SessionInfo Info)> FindSessions()
    {
        var sessions = new List<(object, SessionInfo)>();
        if (!Osascript.IsAvailable()) return sessions;

        var running = Osascript.TryRun(
            "tell application \"System Events\" to (name of processes) contains \"Microsoft Word\"");
        if (running != "true") return sessions;

        int? pid = null;
        var pidRaw = Osascript.TryRun(
            "tell application \"System Events\" to unix id of (first process whose name is \"Microsoft Word\")");
        if (int.TryParse(pidRaw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n)) pid = n;

        string? title = Osascript.TryRun("tell application \"Microsoft Word\" to name of active document");
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

internal sealed class WordMacInfoCommand : IBridgeCommand
{
    public string Name => "info";
    public string Usage => "info   (prints Word version + active document)";

    public Task<int> RunAsync(object comRoot, string[] args, TextWriter output)
    {
        var app = new WdMacApp();
        try { output.WriteLine($"Word version:  {app.Version}"); }
        catch { output.WriteLine("Word version:  (Word not running or unreachable)"); }
        try { output.WriteLine($"Visible:       {app.Visible}"); } catch { }
        try { output.WriteLine($"ActiveDoc:     {app.ActiveDocumentName ?? "(none)"}"); } catch { }
        try { output.WriteLine($"Path:          {app.ActiveDocumentPath ?? "(unsaved)"}"); } catch { }
        try { output.WriteLine($"Documents:     {app.DocumentCount}"); } catch { }
        return Task.FromResult(0);
    }
}

internal sealed class WordMacExtractTextCommand : IBridgeCommand
{
    public string Name => "extract-text";
    public string Usage => "extract-text   (dumps full text of active document)";

    public Task<int> RunAsync(object comRoot, string[] args, TextWriter output)
    {
        var app = new WdMacApp();
        try
        {
            output.Write(app.Content);
            return Task.FromResult(0);
        }
        catch (OsascriptException ex)
        {
            output.WriteLine($"ERROR: {ex.Message}");
            return Task.FromResult(4);
        }
    }
}

internal sealed class WordMacDocStatsCommand : IBridgeCommand
{
    public string Name => "doc-stats";
    public string Usage => "doc-stats   (prints word / paragraph counts)";

    public Task<int> RunAsync(object comRoot, string[] args, TextWriter output)
    {
        var app = new WdMacApp();
        try
        {
            output.WriteLine($"Document:   {app.ActiveDocumentName ?? "(none)"}");
            output.WriteLine($"Words:      {app.WordCount}");
            output.WriteLine($"Paragraphs: {app.ParagraphCount}");
            return Task.FromResult(0);
        }
        catch (OsascriptException ex)
        {
            output.WriteLine($"ERROR: {ex.Message}");
            return Task.FromResult(4);
        }
    }
}
