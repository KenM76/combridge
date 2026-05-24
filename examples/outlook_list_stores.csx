// outlook_list_stores.csx
//
// List every configured Outlook store (mailbox / PST / OST) with account
// info and default-folder item counts. Demonstrates NameSpace.Stores
// iteration + Store properties.
//
// Run:  combridge outlook run-script examples\outlook_list_stores.csx -

Console.WriteLine($"Outlook user: {olNs.CurrentUser?.Name ?? "(unknown)"}");
Console.WriteLine($"Stores: {olNs.Stores.Count}");
Console.WriteLine();

for (int i = 1; i <= olNs.Stores.Count; i++)
{
    Store store = olNs.Stores[i];
    string kind = store.ExchangeStoreType switch
    {
        OlExchangeStoreType.olExchangeMailbox        => "Exchange mailbox",
        OlExchangeStoreType.olExchangePublicFolder   => "Exchange public folder",
        OlExchangeStoreType.olAdditionalExchangeMailbox => "Additional Exchange",
        OlExchangeStoreType.olNotExchange            => "Non-Exchange",
        OlExchangeStoreType.olPrimaryExchangeMailbox => "Primary Exchange",
        _                                            => "Unknown",
    };
    string path = store.FilePath ?? "(server-side)";

    Console.WriteLine($"[{i}] {store.DisplayName}");
    Console.WriteLine($"     kind: {kind}");
    Console.WriteLine($"     file: {path}");

    // Default-folder item counts (best effort — may fail on permissions-restricted shares).
    try
    {
        var inbox = store.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
        Console.WriteLine($"     inbox: {inbox.Items.Count} items");
    }
    catch { Console.WriteLine($"     inbox: (unreadable)"); }

    Console.WriteLine();
}
