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

/// <summary>
/// <c>outlook search</c> for macOS — recursive AppleScript-driven mail
/// search. Same CLI shape as the Windows Outlook plugin's search command
/// so a ScripTree `.scriptree` wrapping it works on both OSes.
/// </summary>
/// <remarks>
/// <para>
/// Architectural parity with Windows:
/// <list type="bullet">
///   <item>Same flag names + meanings: <c>--query</c>, <c>--store</c>
///         (called "account" here but matches by substring same way),
///         <c>--folder</c>, <c>--fields</c>, <c>--max</c>, <c>--since</c>,
///         <c>--snippet</c></item>
///   <item>Same TSV output columns: date, account/store, folder, sender,
///         subject, [snippets]</item>
///   <item>Same Regex snippet extraction (±60 chars × up to 3 windows, whitespace collapsed)</item>
/// </list>
/// </para>
/// <para>
/// Differences from Windows:
/// <list type="bullet">
///   <item>One big osascript invocation walks everything; no per-folder
///         shell-out overhead, but the AppleScript <c>whose</c> filter is
///         significantly slower than DASL <c>Restrict</c></item>
///   <item>No StoreID dedup (Mac Outlook doesn't expose one); duplicates
///         per account name are possible but rare</item>
///   <item>Doesn't work against "New Outlook for Mac" (the 2024+ Catalyst
///         build); only classic Outlook for Mac is supported</item>
/// </list>
/// </para>
/// </remarks>
internal sealed class OlMacSearchCommand : IBridgeCommand
{
    public string Name => "search";
    public string Usage =>
        "search --query \"<text>\" [--store <substr>] [--folder <substr>] " +
        "[--fields subject,body] [--max N] [--since yyyy-MM-dd] [--snippet]";

    public Task<int> RunAsync(object comRoot, string[] args, TextWriter output)
    {
        string? query = null, storeFilter = null, folderFilter = null, fields = "subject,body";
        int max = int.MaxValue;
        DateTime? since = null;
        bool snippet = false;

        for (int i = 0; i < args.Length; i++)
        {
            string a = args[i];
            string? next() => i + 1 < args.Length ? args[++i] : null;
            switch (a)
            {
                case "--query":   query = next(); break;
                case "--store":   storeFilter = next(); break;
                case "--folder":  folderFilter = next(); break;
                case "--fields":  fields = next() ?? fields; break;
                case "--max":     if (int.TryParse(next(), out var m)) max = m; break;
                case "--since":   if (DateTime.TryParse(next(), out var d)) since = d; break;
                case "--snippet": snippet = true; break;
                default:
                    output.WriteLine($"WARN: unknown flag '{a}' ignored.");
                    break;
            }
        }

        if (string.IsNullOrWhiteSpace(query))
        {
            output.WriteLine($"USAGE: {Usage}");
            return Task.FromResult(64);
        }

        var wantSubject = fields.Contains("subject", StringComparison.OrdinalIgnoreCase);
        var wantBody    = fields.Contains("body",    StringComparison.OrdinalIgnoreCase);
        if (!wantSubject && !wantBody)
        {
            output.WriteLine("ERROR: --fields must include at least one of 'subject', 'body'.");
            return Task.FromResult(64);
        }

        var collapseWs = new System.Text.RegularExpressions.Regex(@"\s+",
            System.Text.RegularExpressions.RegexOptions.Compiled);
        var hitRx = new System.Text.RegularExpressions.Regex(
            System.Text.RegularExpressions.Regex.Escape(query),
            System.Text.RegularExpressions.RegexOptions.IgnoreCase |
            System.Text.RegularExpressions.RegexOptions.Compiled);

        output.WriteLine($"# search: query={query} fields={fields}{(since.HasValue ? $" since={since.Value:yyyy-MM-dd}" : "")}{(storeFilter is null ? "" : $" store~={storeFilter}")}");
        output.WriteLine($"# columns: date\\taccount\\tfolder\\tsender\\tsubject{(snippet ? "\\tsnippets" : "")}");

        var app = new OlMacApp();
        var hits = app.Search(
            query: query,
            searchSubject: wantSubject,
            searchBody:    wantBody,
            accountFilter: storeFilter,
            folderFilter:  folderFilter,
            since:         since,
            max:           max,
            wantBody:      snippet);

        int emitted = 0;
        foreach (var h in hits)
        {
            if (emitted >= max) break;
            string date = h.ReceivedTime?.ToString("yyyy-MM-dd HH:mm") ?? "?";
            string row = $"{date}\t{h.Account}\t{h.FolderPath}\t{h.SenderName}\t{h.Subject}";
            if (snippet)
            {
                var snippets = ExtractSnippets(h.Body, hitRx, collapseWs);
                row += $"\t{string.Join(" ¶ ", snippets)}";
            }
            output.WriteLine(row);
            emitted++;
        }

        output.WriteLine();
        output.WriteLine($"# hits emitted:    {emitted}{(emitted >= max && hits.Count > emitted ? " (capped by --max)" : "")}");
        return Task.FromResult(0);
    }

    /// <summary>
    /// Same snippet shape as the Windows plugin's OlSearchCommand —
    /// ±60-char windows, collapsed whitespace, up to 3 non-overlapping.
    /// </summary>
    private static string[] ExtractSnippets(
        string body,
        System.Text.RegularExpressions.Regex hitRx,
        System.Text.RegularExpressions.Regex collapseWs)
    {
        if (string.IsNullOrEmpty(body)) return Array.Empty<string>();
        var flat = collapseWs.Replace(body, " ").Trim();
        var matches = hitRx.Matches(flat);
        if (matches.Count == 0) return Array.Empty<string>();

        var snippets = new List<string>();
        int lastEnd = -1;
        foreach (System.Text.RegularExpressions.Match m in matches)
        {
            int start = Math.Max(0, m.Index - 60);
            int end   = Math.Min(flat.Length, m.Index + m.Length + 60);
            if (start <= lastEnd) continue;
            var window = flat.Substring(start, end - start);
            if (start > 0) window = "…" + window;
            if (end < flat.Length) window += "…";
            snippets.Add(window);
            lastEnd = end;
            if (snippets.Count >= 3) break;
        }
        return snippets.ToArray();
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
