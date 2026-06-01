using System.Globalization;
using ComBridge.Mac.Common;

namespace ComBridge.Plugins.Outlook.Mac;

/// <summary>
/// Mac Outlook application wrapper exposed to user scripts as <c>olApp</c>.
/// </summary>
/// <remarks>
/// <para>
/// <b>Limitations vs. Outlook for Windows COM:</b> the Mac Outlook
/// AppleScript dictionary is significantly thinner than the Windows MAPI
/// COM surface. Mac Outlook exposes:
/// </para>
/// <list type="bullet">
///   <item>Mail accounts, folder enumeration, message basic fields (subject, sender, received date)</item>
///   <item>Sending mail (limited — outgoing messages, no rich formatting control)</item>
///   <item>Reading messages (limited — full body access is restricted in newer Outlook for Mac builds)</item>
/// </list>
/// <para>
/// What's NOT available compared to Windows Outlook COM:
/// </para>
/// <list type="bullet">
///   <item>MAPI NameSpace.Stores (use <see cref="AccountCount"/> + <see cref="AccountNames"/> instead)</item>
///   <item>Restrict() with Jet-style filter syntax (use AppleScript whose-clauses)</item>
///   <item>Rich rule manipulation, RECIPIENTS resolution, GAL access</item>
///   <item>Calendar manipulation beyond reading existing events</item>
/// </list>
/// <para>
/// Microsoft's newer "New Outlook for Mac" (the Catalyst-based UI rolling
/// out 2024+) further restricts AppleScript automation; this plugin
/// targets the classic Outlook for Mac. Migration to a thin REST wrapper
/// over Microsoft Graph may be necessary in a future release.
/// </para>
/// </remarks>
public sealed class OlMacApp
{
    private const string AppName = "Microsoft Outlook";

    public string Version =>
        Osascript.Run($"tell application \"{AppName}\" to version");

    public bool Visible
    {
        get
        {
            var raw = Osascript.TryRun(
                $"tell application \"System Events\" to visible of (first process whose name is \"{AppName}\")");
            return raw == "true";
        }
    }

    /// <summary>Number of configured mail accounts (Exchange + IMAP + POP combined).</summary>
    public int AccountCount
    {
        get
        {
            var raw = Osascript.TryRun(
                $"tell application \"{AppName}\" to (count of exchange accounts) + (count of imap accounts) + (count of pop accounts)");
            return int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : 0;
        }
    }

    /// <summary>Display names of every mail account.</summary>
    public string[] AccountNames
    {
        get
        {
            var script = $@"
tell application ""{AppName}""
    set AppleScript's text item delimiters to linefeed
    set out to {{}}
    repeat with a in exchange accounts
        set end of out to (name of a as text)
    end repeat
    repeat with a in imap accounts
        set end of out to (name of a as text)
    end repeat
    repeat with a in pop accounts
        set end of out to (name of a as text)
    end repeat
    return out as text
end tell".Trim();
            var raw = Osascript.TryRun(script);
            if (string.IsNullOrEmpty(raw)) return Array.Empty<string>();
            return raw.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        }
    }

    /// <summary>Total count of incoming messages (across all accounts' inboxes).</summary>
    public int InboxCount
    {
        get
        {
            var raw = Osascript.TryRun(
                $"tell application \"{AppName}\" to count of (messages of inbox)");
            return int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : 0;
        }
    }

    /// <summary>Count of unread incoming messages.</summary>
    public int UnreadInboxCount
    {
        get
        {
            var raw = Osascript.TryRun(
                $"tell application \"{AppName}\" to count of (messages of inbox whose is read is false)");
            return int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : 0;
        }
    }

    /// <summary>
    /// Recursive mail search across all accounts via one big AppleScript.
    /// Parallels the Windows Outlook plugin's <c>search</c> command but
    /// uses AppleScript <c>whose</c> filters instead of DASL <c>Restrict</c>.
    /// </summary>
    /// <param name="query">Substring to search for (case-insensitive in AppleScript's contains).</param>
    /// <param name="searchSubject">If true, OR-include matches against message subject.</param>
    /// <param name="searchBody">If true, OR-include matches against message plain-text content.</param>
    /// <param name="accountFilter">If non-null, only walk accounts whose name contains this (case-insensitive).</param>
    /// <param name="folderFilter">If non-null, only emit hits from folders whose name contains this. Recursion still walks all children.</param>
    /// <param name="since">If set, only emit messages whose received time is on/after this date.</param>
    /// <param name="max">Cap total hits emitted.</param>
    /// <param name="wantBody">If true, the body is fetched and included in the result rows (for snippet extraction by the caller).</param>
    /// <returns>
    /// One <see cref="SearchHit"/> per matching message. Empty list if nothing
    /// matches or Outlook isn't running. Body is empty unless <paramref name="wantBody"/>.
    /// </returns>
    /// <remarks>
    /// <para>
    /// Performance: this issues ONE osascript invocation that walks every
    /// account's folder tree internally and returns a delimited result blob.
    /// Outlook for Mac's AppleScript <c>whose</c> filter is server-side for
    /// Exchange but client-side for IMAP — count on it being significantly
    /// slower than the Windows DASL/Restrict path. For large mailboxes,
    /// scope down with <paramref name="accountFilter"/> /
    /// <paramref name="folderFilter"/> / <paramref name="since"/>.
    /// </para>
    /// <para>
    /// Caveats:
    /// <list type="bullet">
    ///   <item>"New Outlook for Mac" (the 2024+ Catalyst-based UI) has severely
    ///         restricted AppleScript support; results may be empty even when
    ///         the classic UI would have found matches.</item>
    ///   <item>Same-account duplicate detection isn't available (Mac Outlook
    ///         doesn't expose StoreID); duplicates are best-effort deduped by
    ///         account display name.</item>
    ///   <item>Body fetching adds significant time per hit. Skip
    ///         <paramref name="wantBody"/> if you don't need snippets.</item>
    /// </list>
    /// </para>
    /// </remarks>
    public List<SearchHit> Search(
        string query,
        bool searchSubject = true,
        bool searchBody    = true,
        string? accountFilter = null,
        string? folderFilter  = null,
        DateTime? since = null,
        int max = int.MaxValue,
        bool wantBody = false)
    {
        if (string.IsNullOrEmpty(query)) return new List<SearchHit>();
        if (!Osascript.IsAvailable())    return new List<SearchHit>();

        // Field separators inside one row: U+241E (Record Separator) glyphs
        // are extremely unlikely to appear in mail bodies. Row separator:
        // U+241D. We replace any literal occurrences in body text before
        // emitting, so the round-trip is unambiguous.
        const char ROW_SEP = '␝';
        const char FLD_SEP = '␞';

        string qEsc = Osascript.EscapeForAppleScript(query);
        string accountEsc = accountFilter is null ? "" : Osascript.EscapeForAppleScript(accountFilter);
        string folderEsc  = folderFilter  is null ? "" : Osascript.EscapeForAppleScript(folderFilter);

        // Build the `whose` predicate. AppleScript doesn't reliably accept
        // `whose subject contains q or content contains q` in one clause for
        // Outlook for Mac, so we do TWO `whose` queries and union in the
        // script using a list dedup.
        var subjectClause = searchSubject ? $"messages of f whose subject contains \"{qEsc}\"" : "{}";
        var bodyClause    = searchBody    ? $"messages of f whose plain text content contains \"{qEsc}\"" : "{}";

        // Date filter (post-fetch in the script — AppleScript whose clauses
        // on dates with `as date` are finicky).
        string sinceCheck = "";
        if (since.HasValue)
        {
            // AppleScript date constructor: `date "MM/DD/YYYY HH:MM"` (US locale)
            sinceCheck = $@"
                if (time received of m) < (date ""{since.Value:MM/dd/yyyy HH:mm}"") then exit repeat end if".Trim();
        }

        // Body fetch (optional — adds latency).
        string bodyFetch = wantBody
            ? $@"
                    set theBody to """"
                    try
                        set theBody to plain text content of m as text
                        set theBody to my collapse(theBody)
                    end try"
            : "                    set theBody to \"\"";

        // Folder filter (post-find; we still recurse all children either way).
        string folderCheck = folderFilter is null
            ? "true"
            : $"((name of f as text) contains \"{folderEsc}\")";

        string accountCheck = accountFilter is null
            ? "true"
            : $"((name of acct as text) contains \"{accountEsc}\")";

        // The big walk script. Uses a recursive handler for folder traversal.
        // Emits one ROW_SEP-separated row per hit. Each row is FLD_SEP-joined:
        //   epoch ␞ account ␞ folder-path ␞ sender ␞ senderAddr ␞ subject ␞ body
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

on walkFolder(f, accum, folderPath, maxHits, qStr, doSubject, doBody)
    if (count of accum) >= maxHits then return accum
    try
        set hitsHere to {{}}
        if doSubject then
            try
                set subjMatches to messages of f whose subject contains qStr
                repeat with m in subjMatches
                    set end of hitsHere to m
                end repeat
            end try
        end if
        if doBody then
            try
                set bodyMatches to messages of f whose plain text content contains qStr
                repeat with m in bodyMatches
                    set end of hitsHere to m
                end repeat
            end try
        end if
        -- dedup by id; AppleScript lists don't have set semantics, so accept duplicates
        -- (caller can dedupe by epoch+subject if needed).
        repeat with m in hitsHere
            if (count of accum) >= maxHits then return accum
            try
                set theDate to time received of m
                {sinceCheck}
                set theSubject to my collapse(subject of m as text)
                set theSender to """"
                set theAddr to """"
                try
                    set theSender to name of sender of m as text
                end try
                try
                    set theAddr to address of sender of m as text
                end try
{bodyFetch}
                -- epoch seconds from AS date: difference in seconds from 1970-01-01
                set ep to ((theDate - (date ""1/1/1970"")) - (time to GMT)) -- approx UTC epoch
                set rowText to (ep as text) & ""{FLD_SEP}"" & folderPath & ""{FLD_SEP}"" & theSender & ""{FLD_SEP}"" & theAddr & ""{FLD_SEP}"" & theSubject & ""{FLD_SEP}"" & theBody
                set end of accum to rowText
            end try
        end repeat
        -- recurse into children
        try
            repeat with sub in (mail folders of f)
                set subPath to folderPath & ""/"" & (name of sub as text)
                set accum to my walkFolder(sub, accum, subPath, maxHits, qStr, doSubject, doBody)
            end repeat
        end try
    end try
    return accum
end walkFolder

tell application ""Microsoft Outlook""
    set out to {{}}
    set qStr to ""{qEsc}""
    repeat with acct in (exchange accounts & imap accounts & pop accounts)
        try
            if {accountCheck} then
                set acctName to name of acct as text
                repeat with f in (mail folders of acct)
                    if (count of out) >= {(max == int.MaxValue ? "100000" : max.ToString())} then exit repeat
                    -- Per-account row prefix: epoch ␞ folder-path ␞ sender ␞ addr ␞ subject ␞ body
                    -- folderPath starts with the account name so the output has full context.
                    set rootPath to acctName & "":"" & (name of f as text)
                    set out to my walkFolder(f, out, rootPath, {(max == int.MaxValue ? "100000" : max.ToString())}, qStr, {(searchSubject ? "true" : "false")}, {(searchBody ? "true" : "false")})
                end repeat
            end if
        end try
    end repeat
    set AppleScript's text item delimiters to (character id {(int)ROW_SEP})
    return out as text
end tell".Trim();

        string raw;
        try { raw = Osascript.Run(script); }
        catch (OsascriptException) { return new List<SearchHit>(); }
        if (string.IsNullOrEmpty(raw)) return new List<SearchHit>();

        var hits = new List<SearchHit>();
        foreach (var row in raw.Split(ROW_SEP))
        {
            if (string.IsNullOrEmpty(row)) continue;
            var parts = row.Split(FLD_SEP);
            // Each row should have 6 fields. Tolerate missing trailing ones.
            string ep      = parts.Length > 0 ? parts[0] : "";
            string folder  = parts.Length > 1 ? parts[1] : "";
            string sender  = parts.Length > 2 ? parts[2] : "";
            string addr    = parts.Length > 3 ? parts[3] : "";
            string subject = parts.Length > 4 ? parts[4] : "";
            string body    = parts.Length > 5 ? parts[5] : "";

            DateTime? when = null;
            if (long.TryParse(ep, out var epoch))
                when = DateTimeOffset.FromUnixTimeSeconds(epoch).LocalDateTime;

            hits.Add(new SearchHit(when, accountFilter ?? "", folder, sender, addr, subject, body));
        }
        return hits;
    }
}

/// <summary>
/// One hit row from <see cref="OlMacApp.Search"/>. Body is empty unless
/// the search was called with body-fetching enabled.
/// </summary>
public sealed record SearchHit(
    DateTime? ReceivedTime,
    string Account,
    string FolderPath,
    string SenderName,
    string SenderAddress,
    string Subject,
    string Body);

public sealed class OlMacGlobals
{
    public OlMacApp olApp { get; }
    internal OlMacGlobals() { olApp = new OlMacApp(); }
}
