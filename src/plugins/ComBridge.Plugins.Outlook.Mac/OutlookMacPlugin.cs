using System.Globalization;
using System.Runtime.InteropServices;
using ComBridge.Core;
using ComBridge.Mac.Common;
using Microsoft.CodeAnalysis;

namespace ComBridge.Plugins.Outlook.Mac;

public sealed class OutlookMacPlugin : IComBridgePlugin
{
    public string Name => "outlook";
    public string Description => "Microsoft Outlook for macOS (AppleScript backend, limited dictionary). Globals: olApp.";
    public string[] ProgIds => new[] { "Microsoft Outlook" };
    public bool AllowCreateNew => true;
    public Type GlobalsType => typeof(OlMacGlobals);

    public IReadOnlyCollection<OSPlatform> SupportedPlatforms => new[] { OSPlatform.OSX };

    public object CreateGlobals(object comRoot) => new OlMacGlobals();

    public IEnumerable<MetadataReference> ScriptReferences
    {
        get
        {
            yield return MetadataReference.CreateFromFile(typeof(OutlookMacPlugin).Assembly.Location);
            yield return MetadataReference.CreateFromFile(typeof(Osascript).Assembly.Location);
        }
    }

    public IEnumerable<string> ScriptUsings => new[] { "ComBridge.Plugins.Outlook.Mac" };

    public IEnumerable<IBridgeCommand> Commands => new IBridgeCommand[]
    {
        new OlMacInfoCommand(),
        new OlMacListAccountsCommand(),
        new OlMacSearchCommand(),
        new OlMacGetCommand(),
    };

    public List<(object Root, SessionInfo Info)> FindSessions()
    {
        var sessions = new List<(object, SessionInfo)>();
        if (!Osascript.IsAvailable()) return sessions;

        var running = Osascript.TryRun(
            "tell application \"System Events\" to (name of processes) contains \"Microsoft Outlook\"");
        if (running != "true") return sessions;

        int? pid = null;
        var pidRaw = Osascript.TryRun(
            "tell application \"System Events\" to unix id of (first process whose name is \"Microsoft Outlook\")");
        if (int.TryParse(pidRaw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n)) pid = n;

        // Outlook for Mac doesn't have an obvious "active folder" concept exposed
        // via AppleScript in the same way Windows Outlook does. Use a generic title.
        string? title = "Outlook";

        var desc = pid is int pp ? $"pid={pp}  title={title}" : title!;
        sessions.Add((new object(), new SessionInfo(1, pid, title, desc)));
        return sessions;
    }

    public (int? Pid, string? Title) DescribeInstance(object comRoot) => (null, null);
}

internal sealed class OlMacInfoCommand : IBridgeCommand
{
    public string Name => "info";
    public string Usage => "info   (prints Outlook version + account/inbox summary)";

    public Task<int> RunAsync(object comRoot, string[] args, TextWriter output)
    {
        var app = new OlMacApp();
        try { output.WriteLine($"Outlook version: {app.Version}"); }
        catch { output.WriteLine("Outlook version: (not running)"); }
        try { output.WriteLine($"Visible:         {app.Visible}"); } catch { }
        try { output.WriteLine($"Accounts:        {app.AccountCount}"); } catch { }
        try { output.WriteLine($"Inbox total:     {app.InboxCount}"); } catch { }
        try { output.WriteLine($"Inbox unread:    {app.UnreadInboxCount}"); } catch { }
        return Task.FromResult(0);
    }
}


internal sealed class OlMacListAccountsCommand : IBridgeCommand
{
    public string Name => "list-accounts";
    public string Usage => "list-accounts   (prints display name of every configured account)";

    public Task<int> RunAsync(object comRoot, string[] args, TextWriter output)
    {
        var app = new OlMacApp();
        try
        {
            var names = app.AccountNames;
            for (int i = 0; i < names.Length; i++) output.WriteLine($"  [{i + 1}] {names[i]}");
            output.WriteLine($"\n{names.Length} account(s).");
            return Task.FromResult(0);
        }
        catch (OsascriptException ex)
        {
            output.WriteLine($"ERROR: {ex.Message}");
            return Task.FromResult(4);
        }
    }
}
