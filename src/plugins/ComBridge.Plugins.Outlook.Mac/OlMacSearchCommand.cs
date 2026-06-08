using System.Text;
using System.Text.RegularExpressions;
using ComBridge.Core;
using ComBridge.Mac.Common;

namespace ComBridge.Plugins.Outlook.Mac;

/// <summary>
/// <c>outlook search</c> v2 for macOS — AppleScript-driven multi-term
/// recursive mail search with sender-field support, date windows,
/// word-boundary filtering, and message-id emission. Same CLI shape and
/// output columns as the Windows plugin so a single ScripTree
/// <c>.scriptree</c> works on both OSes.
/// </summary>
/// <remarks>
/// <para>
/// Mac differences from Windows (architecturally):
/// </para>
/// <list type="bullet">
///   <item><b>No DASL / no <c>ci_phrasematch</c>.</b> AppleScript's
///         <c>whose</c> filter is substring-only and is the only
///         pre-filter we have. The word-boundary check therefore always
///         runs C#-side after AppleScript returns the candidate set.
///         There's no "fast path" for indexed stores — every search is
///         the equivalent of the Windows fallback path.</item>
///   <item><b>Per-term separate <c>whose</c> queries.</b> Mac Outlook's
///         dictionary doesn't reliably accept compound
///         <c>whose subject contains "a" or subject contains "b"</c>
///         predicates against messages, so we issue ONE
///         <c>whose</c> per (term × field) and dedupe in-script. Costs
///         <c>terms × fields</c> AppleScript evaluations per folder
///         instead of one — measurable on large mailboxes.</item>
///   <item><b>"New Outlook for Mac" not supported.</b> The 2024+
///         Catalyst UI severely restricts the AppleScript dictionary;
///         results come back empty. Classic Outlook for Mac only.</item>
///   <item><b>Per-item id is an integer</b>, not the opaque hex
///         EntryID Windows uses. We emit it in the <c>entryid</c>
///         column for cross-OS parity, with the account name in the
///         <c>storeid</c> column (since Mac Outlook has no StoreID
///         concept).</item>
/// </list>
/// </remarks>
internal sealed class OlMacSearchCommand : IBridgeCommand
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
                        foreach (var t in qv.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
                            terms.Add(t);
                    break;
                case "--store":   storeFilter = next(); break;
                case "--folder":  folderFilter = next(); break;
                case "--fields":  fields = next() ?? fields; break;
                case "--match":   matchMode = (next() ?? "word").ToLowerInvariant(); break;
                case "--since":   if (DateTime.TryParse(next(), out var d1)) since = d1; break;
                case "--until":   if (DateTime.TryParse(next(), out var d2)) until = d2; break;
                case "--max":     if (int.TryParse(next(), out var m)) max = m; break;
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

        if (!Osascript.IsAvailable())
        {
            output.WriteLine("ERROR: osascript not available.");
            return Task.FromResult(5);
        }

        // Compile per-term regex for attribution + word-boundary post-filter.
        var termRegexes = terms
            .Select(t => (Term: t, Rx: new Regex(
                matchMode == "word" ? $@"\b{Regex.Escape(t)}\b" : Regex.Escape(t),
                RegexOptions.IgnoreCase | RegexOptions.Compiled)))
            .ToList();
        var collapseWs = new Regex(@"\s+", RegexOptions.Compiled);

        output.WriteLine($"# search: query={string.Join(",", terms)} fields={fields} match={matchMode}"
            + (since.HasValue ? $" since={since.Value:yyyy-MM-dd}" : "")
            + (until.HasValue ? $" until={until.Value:yyyy-MM-dd}" : "")
            + (storeFilter is null ? "" : $" store~={storeFilter}")
            + (folderFilter is null ? "" : $" folder~={folderFilter}"));
        output.WriteLine("# columns: date\tstore\tfolder\tsender\tsubject\tmatched\tentryid\tstoreid"
            + (snippet ? "\tsnippets" : ""));

        // Run the AppleScript walker. Returns one delimited row per candidate.
        // We use U+241E (RS) for field separator and U+241D (GS) for row
        // separator — same convention as the v0.4.1 command and extremely
        // unlikely to appear in mail bodies.
        const string FieldSep = "␞";
        const string RowSep   = "␝";

        // Build the per-term × per-field AppleScript whose-clauses. The
        // script collects every (subject contains / content contains /
        // sender contains) hit per folder, dedupes by message id, walks
        // children recursively, returns a delimited blob.
        string accountCheck = storeFilter is null ? "true"
            : $"((name of acct as text) contains \"{Osascript.EscapeForAppleScript(storeFilter)}\")";

        // Each term emits up to 3 whose-clause queries depending on selected fields.
        var queryBlocks = new StringBuilder();
        foreach (var t in terms)
        {
            string esc = Osascript.EscapeForAppleScript(t);
            if (wantSubject)
            {
                queryBlocks.AppendLine($"        try");
                queryBlocks.AppendLine($"            set subjHits to messages of f whose subject contains \"{esc}\"");
                queryBlocks.AppendLine($"            repeat with m in subjHits");
                queryBlocks.AppendLine($"                set end of cand to m");
                queryBlocks.AppendLine($"            end repeat");
                queryBlocks.AppendLine($"        end try");
            }
            if (wantBody)
            {
                queryBlocks.AppendLine($"        try");
                queryBlocks.AppendLine($"            set bodyHits to messages of f whose plain text content contains \"{esc}\"");
                queryBlocks.AppendLine($"            repeat with m in bodyHits");
                queryBlocks.AppendLine($"                set end of cand to m");
                queryBlocks.AppendLine($"            end repeat");
                queryBlocks.AppendLine($"        end try");
            }
            if (wantFrom)
            {
                queryBlocks.AppendLine($"        try");
                queryBlocks.AppendLine($"            set fromHits to messages of f whose sender contains \"{esc}\"");
                queryBlocks.AppendLine($"            repeat with m in fromHits");
                queryBlocks.AppendLine($"                set end of cand to m");
                queryBlocks.AppendLine($"            end repeat");
                queryBlocks.AppendLine($"        end try");
            }
        }

        // Body fetch when snippet is requested OR when word-mode needs the
        // body for re-scan. (For Mac we always include body in the result
        // blob when wanted, since one big script is cheaper than per-message
        // shell-outs.)
        bool needBody = wantBody || snippet;
        string bodyFetch = needBody
            ? "                    try\n                        set theBody to plain text content of m as text\n                        set theBody to my collapse(theBody)\n                    end try"
            : "                    set theBody to \"\"";

        // The full walk script.
        string script = $@"
on collapse(s)
    set t to s
    set AppleScript's text item delimiters to character id 10
    set parts to text items of t
    set AppleScript's text item delimiters to "" ""
    set t to parts as text
    set AppleScript's text item delimiters to character id 13
    set parts to text items of t
    set AppleScript's text item delimiters to "" ""
    set t to parts as text
    set AppleScript's text item delimiters to tab
    set parts to text items of t
    set AppleScript's text item delimiters to "" ""
    set t to parts as text
    return t
end collapse

on walkFolder(f, accum, folderPath, maxHits, seenIds)
    if (count of accum) >= maxHits then return accum
    try
        set cand to {{}}
{queryBlocks}
        repeat with m in cand
            if (count of accum) >= maxHits then return accum
            try
                set msgId to (id of m as text)
                if msgId is not in seenIds then
                    set end of seenIds to msgId
                    set theDate to time received of m
                    set theSubject to my collapse(subject of m as text)
                    set theSender to """"
                    set theAddr to """"
                    try
                        set theSender to name of sender of m as text
                    end try
                    try
                        set theAddr to address of sender of m as text
                    end try
                    set theBody to """"
{bodyFetch}
                    set ep to ((theDate - (date ""1/1/1970"")) - (time to GMT))
                    set rowText to (ep as text) & ""{FieldSep}"" & msgId & ""{FieldSep}"" & folderPath & ""{FieldSep}"" & theSender & ""{FieldSep}"" & theAddr & ""{FieldSep}"" & theSubject & ""{FieldSep}"" & theBody
                    set end of accum to rowText
                end if
            end try
        end repeat
        try
            repeat with sub in (mail folders of f)
                if (count of accum) >= maxHits then return accum
                set subPath to folderPath & ""/"" & (name of sub as text)
                set accum to my walkFolder(sub, accum, subPath, maxHits, seenIds)
            end repeat
        end try
    end try
    return accum
end walkFolder

tell application ""Microsoft Outlook""
    set out to {{}}
    set seenIds to {{}}
    repeat with acct in (exchange accounts & imap accounts & pop accounts)
        try
            if {accountCheck} then
                set acctName to name of acct as text
                repeat with f in (mail folders of acct)
                    if (count of out) >= {(max == int.MaxValue ? "100000" : max.ToString())} then exit repeat
                    set rootPath to acctName & "":"" & (name of f as text)
                    set out to my walkFolder(f, out, rootPath, {(max == int.MaxValue ? "100000" : max.ToString())}, seenIds)
                end repeat
            end if
        end try
    end repeat
    set AppleScript's text item delimiters to (character id {(int)RowSep[0]})
    return out as text
end tell".Trim();

        string raw;
        try { raw = Osascript.Run(script); }
        catch (OsascriptException ex)
        {
            output.WriteLine($"ERROR: AppleScript failed: {ex.Message}");
            return Task.FromResult(5);
        }

        // Parse the delimited blob; apply C#-side date window, folder
        // filter, word-boundary check, attribution.
        int emitted = 0;
        foreach (var row in raw.Split(RowSep))
        {
            if (string.IsNullOrEmpty(row)) continue;
            if (emitted >= max) break;
            var parts = row.Split(FieldSep);
            if (parts.Length < 7) continue;

            // ep id folder sender addr subject body
            DateTime? when = null;
            if (long.TryParse(parts[0], out var epoch))
                when = DateTimeOffset.FromUnixTimeSeconds(epoch).LocalDateTime;
            if (when is null) continue;
            if (since.HasValue && when < since.Value) continue;
            if (until.HasValue && when > until.Value.AddDays(1)) continue;

            string msgId      = parts[1];
            string folderPath = parts[2];
            string senderName = parts[3];
            string senderAddr = parts[4];
            string subject    = parts[5];
            string body       = parts[6];

            // The folder filter is applied per-row (rather than skipping
            // entire folders in AppleScript) because we still want to recurse
            // through non-matching parents to find matching children.
            if (folderFilter is not null &&
                !folderPath.Contains(folderFilter, StringComparison.OrdinalIgnoreCase))
                continue;

            // The account name was prepended to folderPath as `<acct>:<root>/<sub>/...`.
            // Split it back out so the `store` column matches the Windows shape.
            string account = "";
            int colonIdx = folderPath.IndexOf(':');
            if (colonIdx > 0) { account = folderPath[..colonIdx]; folderPath = folderPath[(colonIdx + 1)..]; }

            // Attribution + word-boundary post-filter.
            var matchedTerms = new List<string>();
            foreach (var (term, rx) in termRegexes)
            {
                bool hit = false;
                if (wantSubject && rx.IsMatch(subject)) hit = true;
                if (!hit && wantBody && rx.IsMatch(body)) hit = true;
                if (!hit && wantFrom &&
                    (rx.IsMatch(senderName) || rx.IsMatch(senderAddr))) hit = true;
                if (hit) matchedTerms.Add(term);
            }
            if (matchMode == "word" && matchedTerms.Count == 0) continue;

            string sender = !string.IsNullOrEmpty(senderName) ? senderName
                          : !string.IsNullOrEmpty(senderAddr) ? senderAddr
                          : "?";
            var sb = new StringBuilder();
            sb.Append(when.Value.ToString("yyyy-MM-dd HH:mm")).Append('\t')
              .Append(EscTab(account)).Append('\t')
              .Append(EscTab(folderPath)).Append('\t')
              .Append(EscTab(sender)).Append('\t')
              .Append(EscTab(subject)).Append('\t')
              .Append(string.Join(",", matchedTerms)).Append('\t')
              .Append(msgId).Append('\t')
              .Append(EscTab(account));   // Mac has no StoreID — reuse account name
            if (snippet)
            {
                sb.Append('\t').Append(
                    string.Join(" ¶ ", ExtractSnippets(body, termRegexes.Select(x => x.Rx).ToList(), collapseWs)));
            }
            output.WriteLine(sb.ToString());
            emitted++;
        }

        output.WriteLine();
        output.WriteLine($"# hits emitted: {emitted}");
        return Task.FromResult(0);
    }

    private static string EscTab(string? s)
        => (s ?? "").Replace('\t', ' ').Replace('\n', ' ').Replace('\r', ' ');

    private static string[] ExtractSnippets(string body, List<Regex> termRegexes, Regex collapseWs)
    {
        if (string.IsNullOrEmpty(body) || termRegexes.Count == 0) return Array.Empty<string>();
        var flat = collapseWs.Replace(body, " ").Trim();
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
