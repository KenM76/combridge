using System.Text;
using System.Text.RegularExpressions;
using ComBridge.Core;
using Ol = global::Microsoft.Office.Interop.Outlook;

namespace ComBridge.Plugins.Outlook;

/// <summary>
/// <c>outlook search</c> v2 — recursive mail-content search across MAPI
/// stores with multi-term OR, word-boundary matching, sender-field
/// support, date windows, and always-on EntryID/StoreID output so hits
/// are directly fetchable by <c>outlook get</c>.
/// </summary>
/// <remarks>
/// <para>
/// Replaces the v0.4.0 implementation in full. The original was
/// single-term, substring-only, subject/body-only, since-only, and
/// emitted no EntryID. Every one of those gaps blocked a real
/// "cast a wide net" task documented in
/// <c>FR_outlook_search_v2_multiterm_sender_match_and_get.md</c>.
/// </para>
/// <para>
/// We're the only users (Ken's environment, ScripTreeApps catalog), so
/// backward compatibility is deliberately not preserved — defaults
/// changed where the old default was just wrong (substring → word,
/// subject+body → subject+body+from, optional EntryID → always
/// emitted). Existing wrappers that referenced the old default behavior
/// will need their assumptions updated.
/// </para>
/// <para>
/// Implementation:
/// </para>
/// <list type="bullet">
///   <item><b>Multi-term</b>: <c>--query</c> is repeatable AND accepts
///         comma-separated terms; both forms compose. <c>--query "a,b"
///         --query c</c> → three terms.</item>
///   <item><b>Field surface</b>: <c>--fields</c> accepts
///         <c>subject</c>, <c>body</c>, <c>from</c> (alias <c>sender</c>),
///         in any combination. Default = all three. <c>from</c> maps to
///         both <c>urn:schemas:httpmail:fromname</c> (display name) and
///         <c>urn:schemas:httpmail:fromemail</c> (SMTP address) so
///         "everything from <c>gasspring.ca</c>" works whether the hit
///         comes through the name or the address.</item>
///   <item><b>Match modes</b>: <c>--match</c> accepts <c>word</c>
///         (default) or <c>substring</c>. <c>word</c> uses DASL
///         <c>ci_phrasematch</c> on content-indexed stores — tokenizes
///         on word boundaries so <c>pdac</c> does NOT match inside a
///         base64 blob like <c>ZPDACfM5</c>. On non-indexed stores
///         <c>ci_phrasematch</c> throws "condition is not valid" per
///         folder; we catch per-folder, fall back to the LIKE filter
///         with a C#-side <c>\bterm\b</c> regex pass to drop the same
///         false positives. <c>substring</c> uses LIKE directly and
///         does NOT word-boundary filter — explicit escape hatch for
///         the rare "I really do want the substring".</item>
///   <item><b>Per-hit re-scan</b>: every Restrict result is re-checked
///         in C# against each term × field combination, producing both
///         the <c>matched</c> column (attribution) and the
///         word-boundary drop list (for <c>word</c> mode on the
///         LIKE-fallback path). The old code's
///         <c>totalScanned++; totalHits++</c> shortcut at the Restrict
///         row is gone — Restrict is a pre-filter, not the truth.</item>
///   <item><b>Always-on EntryID/StoreID</b>: every row carries the
///         IDs needed for <c>outlook get --id &lt;EntryID&gt; --store
///         &lt;substr&gt;</c> to re-open the item. EntryIDs are
///         only unique within a store, so both IDs are mandatory.</item>
///   <item><b>Date window</b>: <c>--since</c> (existing) + <c>--until</c>
///         (new) form a closed interval. Either flag alone is open on
///         the other end.</item>
/// </list>
/// </remarks>
internal sealed class OlSearchCommand : IBridgeCommand
{
    public string Name => "search";
    public string Usage =>
        "search <out> [--query \"a,b,c\" ...] [--fields subject,body,from] " +
        "[--match word|substring] [--store <substr>] [--folder <substr>] " +
        "[--since yyyy-MM-dd] [--until yyyy-MM-dd] [--max N] [--snippet]";

    public Task<int> RunAsync(object comRoot, string[] args, TextWriter output)
    {
        var terms = new List<string>();
        string? storeFilter = null, folderFilter = null;
        string fields = "subject,body,from";
        string matchMode = "word";
        int max = int.MaxValue;
        DateTime? since = null, until = null;
        bool snippet = false;

        for (int i = 0; i < args.Length; i++)
        {
            string a = args[i];
            string? next() => i + 1 < args.Length ? args[++i] : null;
            switch (a)
            {
                case "--query":
                    var qv = next();
                    if (!string.IsNullOrEmpty(qv))
                    {
                        // Both repeatable AND comma-separated. One concept,
                        // two surface forms — whichever feels natural at the CLI.
                        foreach (var t in qv.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
                            terms.Add(t);
                    }
                    break;
                case "--store":   storeFilter = next(); break;
                case "--folder":  folderFilter = next(); break;
                case "--fields":  fields = next() ?? fields; break;
                case "--match":   matchMode = (next() ?? "word").ToLowerInvariant(); break;
                case "--since":   if (DateTime.TryParse(next(), out var d1)) since = d1; break;
                case "--until":   if (DateTime.TryParse(next(), out var d2)) until = d2; break;
                case "--max":     if (int.TryParse(next(), out var m))  max  = m;  break;
                case "--snippet": snippet = true; break;
                default:
                    output.WriteLine($"WARN: unknown flag '{a}' ignored.");
                    break;
            }
        }

        if (terms.Count == 0)
        {
            output.WriteLine($"USAGE: {Usage}");
            return Task.FromResult(64);
        }
        if (matchMode != "word" && matchMode != "substring")
        {
            output.WriteLine($"ERROR: --match must be 'word' or 'substring' (got '{matchMode}').");
            return Task.FromResult(64);
        }

        var wantSubject = fields.Contains("subject", StringComparison.OrdinalIgnoreCase);
        var wantBody    = fields.Contains("body",    StringComparison.OrdinalIgnoreCase);
        var wantFrom    = fields.Contains("from",    StringComparison.OrdinalIgnoreCase)
                       || fields.Contains("sender",  StringComparison.OrdinalIgnoreCase);
        if (!wantSubject && !wantBody && !wantFrom)
        {
            output.WriteLine("ERROR: --fields must include at least one of 'subject', 'body', 'from'.");
            return Task.FromResult(64);
        }

        var app = (Ol._Application)comRoot;
        var ns  = app.GetNamespace("MAPI");

        // Compile per-term regexes once. For word mode, anchor with \b
        // (Unicode word boundary on .NET). For substring mode, plain
        // regex on the escaped term — we use it for attribution + snippet
        // windowing, not as a filter (substring lets DASL be the truth).
        var termRegexes = terms
            .Select(t => new
            {
                Term = t,
                Rx = new Regex(
                    matchMode == "word" ? $@"\b{Regex.Escape(t)}\b" : Regex.Escape(t),
                    RegexOptions.IgnoreCase | RegexOptions.Compiled),
            })
            .ToList();
        var collapseWs = new Regex(@"\s+", RegexOptions.Compiled);

        // DASL builder. For ci_phrasematch (word mode) we issue one
        // OR'd clause per (term × field). For LIKE (substring or
        // ci_phrasematch fallback) we issue the same shape with
        // %term% wildcards.
        string BuildDaslOrClauses(bool useCiPhrasematch)
        {
            var clauses = new List<string>();
            foreach (var t in terms)
            {
                var safe = t.Replace("'", "''");
                if (useCiPhrasematch)
                {
                    if (wantSubject) clauses.Add($"\"urn:schemas:httpmail:subject\" ci_phrasematch '{safe}'");
                    if (wantBody)    clauses.Add($"\"urn:schemas:httpmail:textdescription\" ci_phrasematch '{safe}'");
                    if (wantFrom)
                    {
                        clauses.Add($"\"urn:schemas:httpmail:fromname\" ci_phrasematch '{safe}'");
                        clauses.Add($"\"urn:schemas:httpmail:fromemail\" ci_phrasematch '{safe}'");
                    }
                }
                else
                {
                    if (wantSubject) clauses.Add($"\"urn:schemas:httpmail:subject\" LIKE '%{safe}%'");
                    if (wantBody)    clauses.Add($"\"urn:schemas:httpmail:textdescription\" LIKE '%{safe}%'");
                    if (wantFrom)
                    {
                        clauses.Add($"\"urn:schemas:httpmail:fromname\"  LIKE '%{safe}%'");
                        clauses.Add($"\"urn:schemas:httpmail:fromemail\" LIKE '%{safe}%'");
                    }
                }
            }
            return string.Join(" OR ", clauses);
        }

        string BuildDateClause()
        {
            var parts = new List<string>();
            if (since.HasValue) parts.Add($"\"urn:schemas:httpmail:datereceived\" >= '{since.Value:yyyy-MM-dd HH:mm}'");
            if (until.HasValue) parts.Add($"\"urn:schemas:httpmail:datereceived\" <= '{until.Value:yyyy-MM-dd} 23:59'");
            return parts.Count == 0 ? "" : " AND " + string.Join(" AND ", parts);
        }

        string primaryDasl  = "@SQL=(" + BuildDaslOrClauses(useCiPhrasematch: matchMode == "word") + ")" + BuildDateClause();
        string fallbackDasl = "@SQL=(" + BuildDaslOrClauses(useCiPhrasematch: false)                + ")" + BuildDateClause();

        var seenStoreIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int totalScanned = 0, totalHits = 0, foldersWalked = 0, storesWalked = 0, fallbackFolders = 0;

        output.WriteLine($"# search: query={string.Join(",", terms)} fields={fields} match={matchMode}"
            + (since.HasValue ? $" since={since.Value:yyyy-MM-dd}" : "")
            + (until.HasValue ? $" until={until.Value:yyyy-MM-dd}" : "")
            + (storeFilter is null ? "" : $" store~={storeFilter}")
            + (folderFilter is null ? "" : $" folder~={folderFilter}"));
        output.WriteLine("# columns: date\tstore\tfolder\tsender\tsubject\tmatched\tentryid\tstoreid"
            + (snippet ? "\tsnippets" : ""));

        for (int si = 1; si <= ns.Stores.Count; si++)
        {
            Ol.Store store;
            try { store = ns.Stores[si]; } catch { continue; }
            if (storeFilter is not null &&
                !(store.DisplayName?.Contains(storeFilter, StringComparison.OrdinalIgnoreCase) ?? false))
                continue;
            string storeId = "";
            try { storeId = store.StoreID ?? ""; } catch { }
            if (!string.IsNullOrEmpty(storeId) && !seenStoreIds.Add(storeId)) continue;

            storesWalked++;
            Ol.Folder root;
            try { root = (Ol.Folder)store.GetRootFolder(); } catch { continue; }

            WalkFolder(root, store.DisplayName ?? "(unnamed store)", storeId,
                       folderFilter, primaryDasl, fallbackDasl, matchMode,
                       termRegexes.Select(x => (x.Term, x.Rx)).ToList(),
                       wantSubject, wantBody, wantFrom, snippet, max, collapseWs,
                       ref totalScanned, ref totalHits, ref foldersWalked, ref fallbackFolders, output);

            if (totalHits >= max) break;
        }

        output.WriteLine();
        output.WriteLine($"# stores walked:    {storesWalked}");
        output.WriteLine($"# folders walked:   {foldersWalked}");
        output.WriteLine($"# folders fallback: {fallbackFolders} (ci_phrasematch unsupported, LIKE+\\bterm\\b regex used)");
        output.WriteLine($"# items scanned:    {totalScanned}");
        output.WriteLine($"# hits emitted:     {totalHits}{(totalHits >= max ? " (capped by --max)" : "")}");
        return Task.FromResult(0);
    }

    /// <summary>
    /// Walk one folder + its children. Runs the primary DASL (word mode
    /// → <c>ci_phrasematch</c>; substring mode → <c>LIKE</c>). If
    /// <c>ci_phrasematch</c> throws "condition is not valid" (the
    /// non-indexed-store signal), retries the same folder with the
    /// LIKE-flavored fallback DASL and increments
    /// <paramref name="fallbackFolders"/> so the consumer knows which
    /// folders went through the slower path.
    /// </summary>
    private static void WalkFolder(
        Ol.Folder folder, string storeName, string storeId,
        string? folderFilter, string primaryDasl, string fallbackDasl, string matchMode,
        List<(string Term, Regex Rx)> termRegexes,
        bool wantSubject, bool wantBody, bool wantFrom, bool snippet,
        int max, Regex collapseWs,
        ref int totalScanned, ref int totalHits, ref int foldersWalked, ref int fallbackFolders,
        TextWriter output)
    {
        foldersWalked++;

        var folderMatches = folderFilter is null ||
            (folder.Name?.Contains(folderFilter, StringComparison.OrdinalIgnoreCase) ?? false);

        if (folderMatches)
        {
            Ol.Items? filtered = null;
            try
            {
                filtered = folder.Items.Restrict(primaryDasl);
                // Force evaluation — ci_phrasematch can throw lazily on first access.
                _ = filtered.Count;
            }
            catch
            {
                // Likely "The condition is not valid." from ci_phrasematch on a
                // non-content-indexed store. Retry with LIKE; if THAT throws
                // too, the folder is genuinely unsearchable (search folders,
                // Internet Calendars, etc.) and we skip silently.
                if (matchMode == "word")
                {
                    try
                    {
                        filtered = folder.Items.Restrict(fallbackDasl);
                        _ = filtered.Count;
                        fallbackFolders++;
                    }
                    catch { filtered = null; }
                }
                else
                {
                    filtered = null;
                }
            }

            if (filtered is not null)
            {
                for (int i = 1; i <= filtered.Count; i++)
                {
                    if (totalHits >= max) return;
                    object item;
                    try { item = filtered[i]; } catch { continue; }
                    if (item is not Ol.MailItem mail) continue;

                    totalScanned++;

                    string date = "?", senderName = "?", senderEmail = "", subject = "", body = "", entryId = "";
                    try { date        = mail.ReceivedTime.ToString("yyyy-MM-dd HH:mm"); } catch { }
                    try { senderName  = mail.SenderName ?? ""; } catch { }
                    try { senderEmail = mail.SenderEmailAddress ?? ""; } catch { }
                    try { subject     = mail.Subject ?? ""; } catch { }
                    try { entryId     = mail.EntryID ?? ""; } catch { }
                    // Body is fetched if we need it for word-mode re-scan
                    // (always, on word mode) OR for snippet extraction.
                    if (wantBody || snippet)
                    {
                        try { body = mail.Body ?? ""; } catch { body = ""; }
                    }

                    // Per-hit re-scan: for each term, check whether it
                    // actually matched the requested field(s). This serves
                    // two purposes:
                    //   1. attribution (the `matched` column)
                    //   2. drop word-mode false positives that DASL LIKE
                    //      let through but \bterm\b regex rejects
                    var matchedTerms = new List<string>();
                    foreach (var (term, rx) in termRegexes)
                    {
                        bool hit = false;
                        if (wantSubject && rx.IsMatch(subject))          hit = true;
                        if (!hit && wantBody && rx.IsMatch(body))        hit = true;
                        if (!hit && wantFrom &&
                            (rx.IsMatch(senderName) || rx.IsMatch(senderEmail))) hit = true;
                        if (hit) matchedTerms.Add(term);
                    }

                    // Word-mode false-positive drop: if DASL returned this
                    // candidate but no term passed the word-boundary check,
                    // it was a base64-like substring false positive. Skip.
                    // (Substring mode trusts DASL — matchedTerms can be
                    // empty only if the field changed between Restrict
                    // and our read, which is rare; emit anyway with empty
                    // matched column rather than silently lose the row.)
                    if (matchMode == "word" && matchedTerms.Count == 0) continue;

                    totalHits++;

                    string sender = !string.IsNullOrEmpty(senderName) ? senderName
                                  : !string.IsNullOrEmpty(senderEmail) ? senderEmail
                                  : "?";
                    string folderPath = SafeFolderPath(folder);
                    var sb = new StringBuilder();
                    sb.Append(date).Append('\t')
                      .Append(EscTab(storeName)).Append('\t')
                      .Append(EscTab(folderPath)).Append('\t')
                      .Append(EscTab(sender)).Append('\t')
                      .Append(EscTab(subject)).Append('\t')
                      .Append(string.Join(",", matchedTerms)).Append('\t')
                      .Append(entryId).Append('\t')
                      .Append(storeId);
                    if (snippet)
                    {
                        sb.Append('\t').Append(
                            string.Join(" ¶ ", ExtractSnippets(body, termRegexes.Select(x => x.Rx).ToList(), collapseWs)));
                    }
                    output.WriteLine(sb.ToString());
                }
            }
        }

        // Always recurse — nested folders matter even when the parent didn't match.
        try
        {
            foreach (Ol.Folder child in folder.Folders)
            {
                if (totalHits >= max) return;
                WalkFolder(child, storeName, storeId, folderFilter, primaryDasl, fallbackDasl, matchMode,
                           termRegexes, wantSubject, wantBody, wantFrom, snippet, max, collapseWs,
                           ref totalScanned, ref totalHits, ref foldersWalked, ref fallbackFolders, output);
            }
        }
        catch { }
    }

    private static string SafeFolderPath(Ol.Folder f)
    {
        try { return f.FolderPath ?? f.Name ?? "?"; }
        catch { try { return f.Name ?? "?"; } catch { return "?"; } }
    }

    private static string EscTab(string? s)
        => (s ?? "").Replace('\t', ' ').Replace('\n', ' ').Replace('\r', ' ');

    /// <summary>
    /// Extract up to 3 non-overlapping ±60-char windows around hits in
    /// <paramref name="body"/>. Whitespace collapsed to single spaces.
    /// Searches for ANY of the term regexes (so multi-term searches
    /// surface snippets for whichever term matched).
    /// </summary>
    private static string[] ExtractSnippets(string body, List<Regex> termRegexes, Regex collapseWs)
    {
        if (string.IsNullOrEmpty(body) || termRegexes.Count == 0) return Array.Empty<string>();
        var flat = collapseWs.Replace(body, " ").Trim();

        // Merge all term hits into one ordered list of (index, length).
        var hits = new List<(int Index, int Length)>();
        foreach (var rx in termRegexes)
            foreach (Match m in rx.Matches(flat))
                hits.Add((m.Index, m.Length));
        if (hits.Count == 0) return Array.Empty<string>();
        hits.Sort((a, b) => a.Index.CompareTo(b.Index));

        var snippets = new List<string>();
        int lastEnd = -1;
        foreach (var (idx, len) in hits)
        {
            int start = Math.Max(0, idx - 60);
            int end   = Math.Min(flat.Length, idx + len + 60);
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
