using System.Text;
using ComBridge.Core;
using Ol = global::Microsoft.Office.Interop.Outlook;

namespace ComBridge.Plugins.Outlook;

/// <summary>
/// <c>outlook get</c> — dump one (or a few) mail items' full headers
/// and body. Pair with <c>outlook search</c> for the
/// "find candidates → read the one true match" workflow.
/// </summary>
/// <remarks>
/// <para>
/// Two resolution paths:
/// </para>
/// <list type="bullet">
///   <item><b><c>--id &lt;EntryID&gt; [--store &lt;substr&gt;]</c></b> — fast
///         and precise. Uses <c>NameSpace.GetItemFromID(entryID, storeID)</c>.
///         EntryIDs are unique only within a store, so when <c>--store</c>
///         is supplied we resolve the matching store's <c>StoreID</c> and
///         pass it. Without <c>--store</c>, Outlook searches the default
///         store and may miss items in other accounts.</item>
///   <item><b><c>--subject &lt;substr&gt; [--store/--folder &lt;substr&gt;]
///         [--max N]</c></b> — convenience walk for interactive use when
///         you don't have the EntryID. Walks every store / folder like
///         <c>search</c> does but emits FULL items (not snippets). Returns
///         the first <c>--max</c> matches (default 1).</item>
/// </list>
/// <para>
/// Output is a one-block-per-item plain-text dump suitable for human
/// reading or piping into a paginator. Blocks are separated by a
/// <c>=== Item N ===</c> header so multi-item dumps stay legible.
/// </para>
/// <para>
/// Flag composition: pass <c>--html</c> to dump <c>HTMLBody</c> in
/// addition to the plain-text body (useful when the plain-text version
/// is sparse or images carry the content). <c>--headers</c> adds the
/// attachment list (name + size) and the message class — for forensics
/// on quirky items where <c>MessageClass != "IPM.Note"</c>.
/// </para>
/// </remarks>
internal sealed class OlGetCommand : IBridgeCommand
{
    public string Name => "get";
    public string Usage =>
        "get <out> ( --id <EntryID> [--store <substr>] | --subject <substr> " +
        "[--store <substr>] [--folder <substr>] [--max N] ) [--html] [--headers]";

    public Task<int> RunAsync(object comRoot, string[] args, TextWriter output)
    {
        string? entryId = null, storeFilter = null, subjectFilter = null, folderFilter = null;
        int max = 1;
        bool html = false, headers = false;

        for (int i = 0; i < args.Length; i++)
        {
            string a = args[i];
            string? next() => i + 1 < args.Length ? args[++i] : null;
            switch (a)
            {
                case "--id":      entryId = next(); break;
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

        if (string.IsNullOrEmpty(entryId) && string.IsNullOrEmpty(subjectFilter))
        {
            output.WriteLine($"USAGE: {Usage}");
            return Task.FromResult(64);
        }
        if (!string.IsNullOrEmpty(entryId) && !string.IsNullOrEmpty(subjectFilter))
        {
            output.WriteLine("ERROR: --id and --subject are mutually exclusive.");
            return Task.FromResult(64);
        }

        var app = (Ol._Application)comRoot;
        var ns  = app.GetNamespace("MAPI");

        // Path 1: direct lookup by EntryID. Resolve StoreID via --store first
        // (substring match on store DisplayName). Falls through to default
        // store if --store is omitted.
        if (!string.IsNullOrEmpty(entryId))
        {
            string? storeId = null;
            if (!string.IsNullOrEmpty(storeFilter))
            {
                for (int si = 1; si <= ns.Stores.Count; si++)
                {
                    Ol.Store s;
                    try { s = ns.Stores[si]; } catch { continue; }
                    if (s.DisplayName?.Contains(storeFilter, StringComparison.OrdinalIgnoreCase) ?? false)
                    {
                        try { storeId = s.StoreID; } catch { }
                        break;
                    }
                }
                if (storeId is null)
                {
                    output.WriteLine($"ERROR: no store matched --store '{storeFilter}'.");
                    return Task.FromResult(2);
                }
            }

            object? item;
            try
            {
                item = storeId is null
                    ? ns.GetItemFromID(entryId)
                    : ns.GetItemFromID(entryId, storeId);
            }
            catch (Exception ex)
            {
                output.WriteLine($"ERROR: GetItemFromID failed: {ex.Message}");
                return Task.FromResult(3);
            }
            if (item is not Ol.MailItem mail)
            {
                output.WriteLine($"ERROR: item is not a MailItem (got '{item?.GetType().Name ?? "null"}').");
                return Task.FromResult(3);
            }

            DumpItem(mail, 1, html, headers, output);
            return Task.FromResult(0);
        }

        // Path 2: subject substring walk. Emit the first --max full matches.
        // Same store dedup + folder traversal as search; per-folder try/catch
        // tolerates unsearchable folders.
        int emitted = 0;
        var seenStoreIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (int si = 1; si <= ns.Stores.Count; si++)
        {
            if (emitted >= max) break;
            Ol.Store store;
            try { store = ns.Stores[si]; } catch { continue; }
            if (storeFilter is not null &&
                !(store.DisplayName?.Contains(storeFilter, StringComparison.OrdinalIgnoreCase) ?? false))
                continue;
            string storeId = "";
            try { storeId = store.StoreID ?? ""; } catch { }
            if (!string.IsNullOrEmpty(storeId) && !seenStoreIds.Add(storeId)) continue;

            Ol.Folder root;
            try { root = (Ol.Folder)store.GetRootFolder(); } catch { continue; }
            WalkForSubject(root, subjectFilter!, folderFilter, max, html, headers,
                           ref emitted, output);
        }

        output.WriteLine();
        output.WriteLine($"# items emitted: {emitted}{(emitted >= max ? " (capped by --max)" : "")}");
        return Task.FromResult(emitted > 0 ? 0 : 4);
    }

    /// <summary>
    /// Walk one folder + children looking for MailItems whose Subject
    /// contains <paramref name="subjectFilter"/>. DumpItem each match
    /// until <paramref name="max"/> reached. Uses Restrict for the
    /// subject pre-filter so we don't iterate every item by hand.
    /// </summary>
    private static void WalkForSubject(
        Ol.Folder folder, string subjectFilter, string? folderFilter, int max,
        bool html, bool headers, ref int emitted, TextWriter output)
    {
        if (emitted >= max) return;

        var folderMatches = folderFilter is null ||
            (folder.Name?.Contains(folderFilter, StringComparison.OrdinalIgnoreCase) ?? false);

        if (folderMatches)
        {
            try
            {
                var qSafe = subjectFilter.Replace("'", "''");
                var dasl  = $"@SQL=(\"urn:schemas:httpmail:subject\" LIKE '%{qSafe}%')";
                var filtered = folder.Items.Restrict(dasl);
                for (int i = 1; i <= filtered.Count; i++)
                {
                    if (emitted >= max) return;
                    object item;
                    try { item = filtered[i]; } catch { continue; }
                    if (item is not Ol.MailItem mail) continue;
                    DumpItem(mail, emitted + 1, html, headers, output);
                    emitted++;
                }
            }
            catch { /* tolerate unsearchable folders */ }
        }

        try
        {
            foreach (Ol.Folder child in folder.Folders)
            {
                if (emitted >= max) return;
                WalkForSubject(child, subjectFilter, folderFilter, max, html, headers, ref emitted, output);
            }
        }
        catch { }
    }

    /// <summary>
    /// Format one MailItem as a labeled-header block followed by the body.
    /// Each property read individually-try/catch'd because COM RPC faults
    /// on any single property shouldn't lose the rest.
    /// </summary>
    private static void DumpItem(Ol.MailItem mail, int index, bool html, bool headers, TextWriter output)
    {
        output.WriteLine($"=== Item {index} ===");
        WriteHeader(output, "Subject",      () => mail.Subject);
        WriteHeader(output, "From",         () =>
            $"{mail.SenderName ?? ""} <{mail.SenderEmailAddress ?? ""}>".Trim());
        WriteHeader(output, "To",           () => mail.To);
        WriteHeader(output, "Cc",           () => mail.CC);
        WriteHeader(output, "ReceivedTime", () => mail.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss"));
        WriteHeader(output, "SentOn",       () => mail.SentOn.ToString("yyyy-MM-dd HH:mm:ss"));
        WriteHeader(output, "Size",         () => $"{mail.Size} bytes");
        WriteHeader(output, "Folder",       () => GetParentFolderPath(mail));
        WriteHeader(output, "EntryID",      () => mail.EntryID);
        WriteHeader(output, "StoreID",      () => mail.Parent is Ol.Folder pf ? pf.StoreID : "");

        if (headers)
        {
            WriteHeader(output, "MessageClass", () => mail.MessageClass);
            try
            {
                var atts = mail.Attachments;
                if (atts is not null && atts.Count > 0)
                {
                    output.WriteLine("Attachments:");
                    for (int ai = 1; ai <= atts.Count; ai++)
                    {
                        try
                        {
                            var att = atts[ai];
                            output.WriteLine($"  - {att.FileName} ({att.Size:N0} bytes)");
                        }
                        catch { }
                    }
                }
            }
            catch { }
        }

        output.WriteLine();
        output.WriteLine("[BODY]");
        try { output.WriteLine(mail.Body ?? "(no plain-text body)"); }
        catch (Exception ex) { output.WriteLine($"(body read failed: {ex.Message})"); }

        if (html)
        {
            output.WriteLine();
            output.WriteLine("[HTMLBODY]");
            try { output.WriteLine(mail.HTMLBody ?? "(no HTML body)"); }
            catch (Exception ex) { output.WriteLine($"(HTML body read failed: {ex.Message})"); }
        }

        output.WriteLine();
    }

    private static void WriteHeader(TextWriter output, string label, Func<string?> read)
    {
        try
        {
            var v = read();
            if (!string.IsNullOrEmpty(v)) output.WriteLine($"{label,-13} {v}");
        }
        catch { /* skip missing/unreadable headers */ }
    }

    private static string GetParentFolderPath(Ol.MailItem mail)
    {
        try
        {
            if (mail.Parent is Ol.Folder f)
                return f.FolderPath ?? f.Name ?? "?";
        }
        catch { }
        return "?";
    }
}
