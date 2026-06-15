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

}

public sealed class OlMacGlobals : ComBridge.Core.IScriptContext
{
    public OlMacApp olApp { get; }

    // Host-injected I/O channels — see ComBridge.Core.IScriptContext.
    public string[] ScriptArgs { get; set; } = Array.Empty<string>();
    public string Stdin { get; set; } = "";

    internal OlMacGlobals() { olApp = new OlMacApp(); }
}
