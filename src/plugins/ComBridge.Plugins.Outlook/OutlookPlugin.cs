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
        new OlSearchCommand(),
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

/// <summary>
/// <c>outlook search</c> — recursive mail-content search across MAPI stores
/// with DASL <c>Restrict</c> for speed. Replaces the "every search is a
/// bespoke .csx" pattern. See FR
/// <c>FR_scripting_dx_and_outlook_search.md</c> § Item 3 for design notes.
/// </summary>
/// <remarks>
/// <para>
/// Implementation notes:
/// <list type="bullet">
///   <item>Builds a DASL filter with <c>@SQL=</c> syntax. The two field options
///         (<c>subject</c>, <c>body</c>) map to the <c>urn:schemas:httpmail:subject</c>
///         and <c>urn:schemas:httpmail:textdescription</c> properties — content-indexed
///         on Exchange/IMAP, dramatically faster than C#-side string-matching.</item>
///   <item>Per-folder try/catch so unscriptable stores (Internet Calendars,
///         search folders) degrade rather than abort the walk.</item>
///   <item>De-dupes stores by <c>StoreID</c> — some mailboxes (e.g. when an
///         account is added twice with different display names) list the same
///         store twice in <c>NameSpace.Stores</c>.</item>
///   <item>Snippet extraction uses <c>System.Text.RegularExpressions</c>
///         (available in default refs as of v0.4.0).</item>
/// </list>
/// </para>
/// <para>
/// Not implemented on Mac Outlook — DASL <c>Restrict</c> isn't available
/// via AppleScript. A Mac equivalent would have to iterate
/// <c>messages of inbox whose ...</c> with whose-clause filters, much slower.
/// </para>
/// </remarks>
internal sealed class OlSearchCommand : IBridgeCommand
{
    public string Name => "search";
    public string Usage =>
        "search --query \"<text>\" [--store <substr>] [--folder <substr>] " +
        "[--fields subject,body] [--max N] [--since yyyy-MM-dd] [--snippet]";

    public Task<int> RunAsync(object comRoot, string[] args, TextWriter output)
    {
        // Parse flags. Conservative — no third-party CLI lib, just match by
        // name and consume the next arg.
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
                case "--max":
                    if (int.TryParse(next(), out var m)) max = m;
                    break;
                case "--since":
                    if (DateTime.TryParse(next(), out var d)) since = d;
                    break;
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

        var app = (Ol._Application)comRoot;
        var ns = app.GetNamespace("MAPI");

        // Build the DASL filter once. Always quote the query for the LIKE
        // clauses; DASL handles single quotes via doubling.
        var qSafe = query.Replace("'", "''");
        var wantSubject = fields.Contains("subject", StringComparison.OrdinalIgnoreCase);
        var wantBody    = fields.Contains("body",    StringComparison.OrdinalIgnoreCase);
        if (!wantSubject && !wantBody)
        {
            output.WriteLine("ERROR: --fields must include at least one of 'subject', 'body'.");
            return Task.FromResult(64);
        }
        var likeClauses = new List<string>();
        if (wantSubject) likeClauses.Add($"\"urn:schemas:httpmail:subject\" LIKE '%{qSafe}%'");
        if (wantBody)    likeClauses.Add($"\"urn:schemas:httpmail:textdescription\" LIKE '%{qSafe}%'");
        var dasl = $"@SQL=({string.Join(" OR ", likeClauses)})";
        if (since.HasValue)
        {
            dasl = $"@SQL=({string.Join(" OR ", likeClauses)}) AND " +
                   $"\"urn:schemas:httpmail:datereceived\" >= '{since.Value:yyyy-MM-dd HH:mm}'";
        }

        // Snippet extraction regex: collapse whitespace then window ±60 chars around the match.
        var collapseWs = new System.Text.RegularExpressions.Regex(@"\s+", System.Text.RegularExpressions.RegexOptions.Compiled);
        var hitRx = new System.Text.RegularExpressions.Regex(
            System.Text.RegularExpressions.Regex.Escape(query),
            System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);

        int totalScanned = 0, totalHits = 0, foldersWalked = 0, storesWalked = 0;
        var seenStoreIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        output.WriteLine($"# search: query={query} fields={fields}{(since.HasValue ? $" since={since.Value:yyyy-MM-dd}" : "")}{(storeFilter is null ? "" : $" store~={storeFilter}")}");
        output.WriteLine($"# columns: date\\tstore\\tfolder\\tsender\\tsubject{(snippet ? "\\tsnippets" : "")}");

        for (int si = 1; si <= ns.Stores.Count; si++)
        {
            Ol.Store store;
            try { store = ns.Stores[si]; } catch { continue; }
            if (storeFilter is not null &&
                !(store.DisplayName?.Contains(storeFilter, StringComparison.OrdinalIgnoreCase) ?? false))
                continue;
            if (!string.IsNullOrEmpty(store.StoreID) && !seenStoreIds.Add(store.StoreID))
                continue; // dedupe: same store ID listed twice

            storesWalked++;
            Ol.Folder root;
            try { root = (Ol.Folder)store.GetRootFolder(); } catch { continue; }

            WalkFolder(root, store.DisplayName ?? "(unnamed store)", folderFilter,
                       dasl, hitRx, collapseWs, snippet, max,
                       ref totalScanned, ref totalHits, ref foldersWalked, output);

            if (totalHits >= max) break;
        }

        output.WriteLine();
        output.WriteLine($"# stores walked:   {storesWalked}");
        output.WriteLine($"# folders walked:  {foldersWalked}");
        output.WriteLine($"# items scanned:   {totalScanned}");
        output.WriteLine($"# hits emitted:    {totalHits}{(totalHits >= max ? " (capped by --max)" : "")}");
        return Task.FromResult(0);
    }

    private static void WalkFolder(
        Ol.Folder folder, string storeName, string? folderFilter,
        string dasl, System.Text.RegularExpressions.Regex hitRx,
        System.Text.RegularExpressions.Regex collapseWs, bool snippet, int max,
        ref int totalScanned, ref int totalHits, ref int foldersWalked, TextWriter output)
    {
        foldersWalked++;

        // Folder filter: if set, only emit hits from folders whose name contains
        // the substring. Always recurse children so nested folders still get
        // walked when the parent doesn't match.
        var folderMatches = folderFilter is null ||
            (folder.Name?.Contains(folderFilter, StringComparison.OrdinalIgnoreCase) ?? false);

        if (folderMatches)
        {
            try
            {
                var filtered = folder.Items.Restrict(dasl);
                for (int i = 1; i <= filtered.Count; i++)
                {
                    if (totalHits >= max) return;
                    object item;
                    try { item = filtered[i]; } catch { continue; }
                    if (item is not Ol.MailItem mail) continue;

                    totalScanned++;
                    totalHits++;

                    string date = "?", sender = "?", subject = "?", body = "";
                    try { date    = mail.ReceivedTime.ToString("yyyy-MM-dd HH:mm"); } catch { }
                    try { sender  = mail.SenderName ?? mail.SenderEmailAddress ?? "?"; } catch { }
                    try { subject = mail.Subject ?? ""; } catch { }
                    if (snippet)
                    {
                        try { body = mail.Body ?? ""; } catch { body = ""; }
                    }

                    string folderPath = SafeFolderPath(folder);
                    string row = $"{date}\t{storeName}\t{folderPath}\t{sender}\t{subject}";
                    if (snippet)
                    {
                        var snippets = ExtractSnippets(body, hitRx, collapseWs);
                        row += $"\t{string.Join(" ¶ ", snippets)}";
                    }
                    output.WriteLine(row);
                }
            }
            catch
            {
                // Some folders (search folders, "Internet Calendars", recipient
                // cache, etc.) throw on Restrict. Skip silently — the walk
                // should degrade, not abort.
            }
        }

        // Recurse into child folders even if THIS folder didn't match (nested
        // structure is normal).
        try
        {
            foreach (Ol.Folder child in folder.Folders)
            {
                if (totalHits >= max) return;
                WalkFolder(child, storeName, folderFilter, dasl, hitRx, collapseWs,
                           snippet, max, ref totalScanned, ref totalHits, ref foldersWalked, output);
            }
        }
        catch { /* tolerate iteration failures */ }
    }

    private static string SafeFolderPath(Ol.Folder f)
    {
        try { return f.FolderPath ?? f.Name ?? "?"; }
        catch { try { return f.Name ?? "?"; } catch { return "?"; } }
    }

    /// <summary>
    /// Extract up to 3 non-overlapping ±60-char windows around hits in
    /// <paramref name="body"/>. Whitespace collapsed to single spaces.
    /// Returns empty array if no hits in the body.
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
            if (start <= lastEnd) continue;   // overlap with previous snippet — skip
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
