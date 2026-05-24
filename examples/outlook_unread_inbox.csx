// outlook_unread_inbox.csx
//
// List unread mail items in the default Inbox. Demonstrates NameSpace folder
// access + Items.Restrict for filtering.
//
// Run:  combridge outlook run-script examples\outlook_unread_inbox.csx -

const int maxShow = 50;

var inbox = olNs.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
Console.WriteLine($"Inbox total items: {inbox.Items.Count}");

// Restrict uses Jet-style query syntax. "[Unread]" is a property accessor.
var unread = inbox.Items.Restrict("[Unread] = true");
unread.Sort("[ReceivedTime]", true /*Descending*/);

int n = 0;
foreach (object item in unread)
{
    if (item is MailItem m)
    {
        n++;
        string from = m.SenderEmailAddress ?? m.SenderName ?? "(unknown)";
        Console.WriteLine($"  [{n,2}] {m.ReceivedTime:yyyy-MM-dd HH:mm}  {from,-40}  → {m.Subject}");
        if (n >= maxShow) { Console.WriteLine($"  ... (capped at {maxShow})"); break; }
    }
}

Console.WriteLine($"\nShown: {n} unread.");
