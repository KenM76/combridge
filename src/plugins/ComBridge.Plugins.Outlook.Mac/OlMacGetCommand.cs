using ComBridge.Core;
using ComBridge.Mac.Common;

namespace ComBridge.Plugins.Outlook.Mac;

/// <summary>
/// <c>outlook get</c> for macOS — dump one (or a few) mail items' full
/// headers and body. Pairs with <c>outlook search</c> for the
/// "find candidates → read the winner" workflow. Same CLI shape as the
/// Windows version so a single ScripTree wrapper works cross-OS.
/// </summary>
/// <remarks>
/// <para>
/// Two resolution paths, mirroring the Windows command:
/// </para>
/// <list type="bullet">
///   <item><b><c>--id &lt;message-id&gt; [--store &lt;substr&gt;]</c></b> —
///         on Mac the id is the integer AppleScript exposes via
///         <c>id of message</c>. Resolved with the AppleScript
///         <c>messages whose id is N</c> filter across the account(s).
///         <c>--store</c> scopes the search to one account name.</item>
///   <item><b><c>--subject &lt;substr&gt; [--store/--folder &lt;substr&gt;]
///         [--max N]</c></b> — recursive walk via AppleScript whose-clause.</item>
/// </list>
/// <para>
/// Output format matches the Windows version: <c>=== Item N ===</c>
/// header per item, labeled headers, then <c>[BODY]</c> with the
/// plain-text content. <c>--html</c> additionally dumps
/// <c>content</c> (Mac Outlook's HTML accessor). <c>--headers</c>
/// adds the attachments list and message class equivalent.
/// </para>
/// </remarks>
internal sealed class OlMacGetCommand : IBridgeCommand
{
    public string Name => "get";
    public string Usage =>
        "get <out> ( --id <id> [--store <substr>] | --subject <substr> " +
        "[--store <substr>] [--folder <substr>] [--max N] ) [--html] [--headers]";

    public Task<int> RunAsync(object comRoot, string[] args, TextWriter output)
    {
        string? idStr = null, storeFilter = null, subjectFilter = null, folderFilter = null;
        int max = 1;
        bool html = false, headers = false;

        for (int i = 0; i < args.Length; i++)
        {
            string a = args[i];
            string? next() => i + 1 < args.Length ? args[++i] : null;
            switch (a)
            {
                case "--id":      idStr = next(); break;
                case "--store":   storeFilter = next(); break;
                case "--subject": subjectFilter = next(); break;
                case "--folder":  folderFilter = next(); break;
                case "--max":     if (int.TryParse(next(), out var m)) max = m; break;
                case "--html":    html = true; break;
                case "--headers": headers = true; break;
                default:
                    output.WriteLine($"WARN: unknown flag '{a}' ignored.");
                    break;
            }
        }

        if (string.IsNullOrEmpty(idStr) && string.IsNullOrEmpty(subjectFilter))
        {
            output.WriteLine($"USAGE: {Usage}");
            return Task.FromResult(64);
        }
        if (!string.IsNullOrEmpty(idStr) && !string.IsNullOrEmpty(subjectFilter))
        {
            output.WriteLine("ERROR: --id and --subject are mutually exclusive.");
            return Task.FromResult(64);
        }
        if (!Osascript.IsAvailable())
        {
            output.WriteLine("ERROR: osascript not available.");
            return Task.FromResult(5);
        }

        // Both paths converge on a single AppleScript that emits N items.
        // The walking script differs by predicate, but the output format and
        // C#-side parsing are identical.
        const string FieldSep = "␞";
        const string RowSep   = "␝";

        string predicate;
        if (!string.IsNullOrEmpty(idStr))
        {
            // Walk accounts/folders looking for `message id <N>`. We have to
            // walk because Mac Outlook's `message id N` global lookup is
            // unreliable in classic Outlook — safer to filter per folder.
            if (!long.TryParse(idStr, out var idNum))
            {
                output.WriteLine($"ERROR: --id must be an integer on macOS (got '{idStr}').");
                return Task.FromResult(64);
            }
            predicate = $"messages of f whose id is {idNum}";
        }
        else
        {
            string esc = Osascript.EscapeForAppleScript(subjectFilter!);
            predicate = $"messages of f whose subject contains \"{esc}\"";
        }

        string accountCheck = storeFilter is null ? "true"
            : $"((name of acct as text) contains \"{Osascript.EscapeForAppleScript(storeFilter)}\")";
        string folderCheck = folderFilter is null ? "true"
            : $"((name of f as text) contains \"{Osascript.EscapeForAppleScript(folderFilter)}\")";

        // Per-item we collect: subject, sender name+addr, recipients, dates,
        // folder path, id, body (always), html content (if --html), message
        // class (if --headers), attachments (if --headers).
        string htmlFetch = html
            ? "                set theHtml to \"\"\n                try\n                    set theHtml to content of m as text\n                end try\n"
            : "                set theHtml to \"\"\n";
        string attsFetch = headers
            ? @"                set attList to """"
                try
                    repeat with a in mail attachments of m
                        try
                            set attList to attList & ""  - "" & (name of a as text) & "" ("" & (file size of a as text) & "" bytes)"" & character id 10
                        end try
                    end repeat
                end try
"
            : "                set attList to \"\"\n";

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
    return t
end collapse

on walkFolder(f, accum, folderPath, maxHits)
    if (count of accum) >= maxHits then return accum
    try
        if not ({folderCheck}) then
            -- still recurse for nested matches
            try
                repeat with sub in (mail folders of f)
                    if (count of accum) >= maxHits then return accum
                    set subPath to folderPath & ""/"" & (name of sub as text)
                    set accum to my walkFolder(sub, accum, subPath, maxHits)
                end repeat
            end try
            return accum
        end if
        try
            set hits to {predicate}
            repeat with m in hits
                if (count of accum) >= maxHits then return accum
                try
                    set theSubject to my collapse(subject of m as text)
                    set theSender to """"
                    set theAddr to """"
                    try
                        set theSender to name of sender of m as text
                    end try
                    try
                        set theAddr to address of sender of m as text
                    end try
                    set theDate to (time received of m as text)
                    set msgId to (id of m as text)
                    set msgSize to """"
                    try
                        set msgSize to (size of m as text)
                    end try
                    set msgClass to """"
                    try
                        set msgClass to (class of m as text)
                    end try
                    set theBody to """"
                    try
                        set theBody to plain text content of m as text
                    end try
{htmlFetch}{attsFetch}
                    set rowText to msgId & ""{FieldSep}"" & folderPath & ""{FieldSep}"" & theSender & ""{FieldSep}"" & theAddr & ""{FieldSep}"" & theSubject & ""{FieldSep}"" & theDate & ""{FieldSep}"" & msgSize & ""{FieldSep}"" & msgClass & ""{FieldSep}"" & theBody & ""{FieldSep}"" & theHtml & ""{FieldSep}"" & attList
                    set end of accum to rowText
                end try
            end repeat
        end try
        try
            repeat with sub in (mail folders of f)
                if (count of accum) >= maxHits then return accum
                set subPath to folderPath & ""/"" & (name of sub as text)
                set accum to my walkFolder(sub, accum, subPath, maxHits)
            end repeat
        end try
    end try
    return accum
end walkFolder

tell application ""Microsoft Outlook""
    set out to {{}}
    repeat with acct in (exchange accounts & imap accounts & pop accounts)
        try
            if {accountCheck} then
                set acctName to name of acct as text
                repeat with f in (mail folders of acct)
                    if (count of out) >= {max} then exit repeat
                    set rootPath to acctName & "":"" & (name of f as text)
                    set out to my walkFolder(f, out, rootPath, {max})
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
        if (string.IsNullOrEmpty(raw))
        {
            output.WriteLine($"# items emitted: 0");
            return Task.FromResult(4);
        }

        int emitted = 0;
        foreach (var row in raw.Split(RowSep))
        {
            if (string.IsNullOrEmpty(row)) continue;
            if (emitted >= max) break;
            var parts = row.Split(FieldSep);
            if (parts.Length < 11) continue;

            string msgId = parts[0], folder = parts[1], sender = parts[2], addr = parts[3],
                   subject = parts[4], date = parts[5], size = parts[6], cls = parts[7],
                   body = parts[8], htmlBody = parts[9], atts = parts[10];

            output.WriteLine($"=== Item {emitted + 1} ===");
            output.WriteLine($"Subject       {subject}");
            output.WriteLine($"From          {sender} <{addr}>");
            output.WriteLine($"ReceivedTime  {date}");
            if (!string.IsNullOrEmpty(size)) output.WriteLine($"Size          {size} bytes");
            output.WriteLine($"Folder        {folder}");
            output.WriteLine($"EntryID       {msgId}");
            if (headers)
            {
                if (!string.IsNullOrEmpty(cls)) output.WriteLine($"MessageClass  {cls}");
                if (!string.IsNullOrEmpty(atts))
                {
                    output.WriteLine("Attachments:");
                    output.WriteLine(atts.TrimEnd());
                }
            }

            output.WriteLine();
            output.WriteLine("[BODY]");
            output.WriteLine(string.IsNullOrEmpty(body) ? "(no plain-text body)" : body);

            if (html)
            {
                output.WriteLine();
                output.WriteLine("[HTMLBODY]");
                output.WriteLine(string.IsNullOrEmpty(htmlBody) ? "(no HTML body)" : htmlBody);
            }
            output.WriteLine();
            emitted++;
        }

        output.WriteLine($"# items emitted: {emitted}");
        return Task.FromResult(emitted > 0 ? 0 : 4);
    }
}
