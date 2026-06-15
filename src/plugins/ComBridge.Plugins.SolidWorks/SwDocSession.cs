using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace ComBridge.Plugins.SolidWorks;

/// <summary>
/// Shared helpers for the typed read-commands (<c>list-configs</c>,
/// <c>list-components</c>) that may need to open a SOLIDWORKS file
/// just to inspect it. Encapsulates the "if already open in this
/// session, reuse it; otherwise open silently read-only, and close
/// afterwards" dance so each command doesn't re-derive the lifetime
/// rules.
/// </summary>
/// <remarks>
/// <para>
/// Safety discipline baked in (from personal_rag):
/// </para>
/// <list type="bullet">
///   <item><b>OpenDoc6 silent-config-arg bug</b>
///         (<c>lesson_20260512_opendoc6_config_arg_silent_bug.md</c>):
///         passing <c>""</c> as the config arg loads whatever config
///         the file was LAST SAVED in, NOT "the default." Callers
///         that want a specific config MUST pass it through to
///         <see cref="OpenForReadAsync"/>, AND defense-in-depth verify
///         the active config matches afterwards (<c>ShowConfiguration2</c>
///         + <c>ForceRebuild3</c> if not). The verification step lives
///         in <see cref="ListComponentsCommand"/> because it's the only
///         command that uses a config arg today.</item>
///   <item><b>CloseDoc leaves components resident</b>
///         (<c>lesson_20260608_closedoc_leaves_components_resident.md</c>):
///         <see cref="ISldWorks.CloseDoc"/> only closes the named
///         top-level document. Component docs SW loaded as references
///         remain in the session. We do NOT call
///         <c>CloseAllDocuments(true)</c> as a cleanup because that
///         would close the USER'S work with unsaved changes destroyed.
///         The accumulation is a tradeoff documented in each command's
///         XML — users running these many times in succession against
///         different assemblies will see component-doc growth.</item>
///   <item><b>Active-doc fallback</b>
///         (<c>lesson_20260427_active_doc_fallback_pattern.md</c>):
///         when a command takes a path and no path is given, fall back
///         to <c>swApp.ActiveDoc</c>. Convention: empty/missing path
///         input == "use the active doc."</item>
/// </list>
/// </remarks>
internal static class SwDocSession
{
    /// <summary>
    /// Acquire a <see cref="IModelDoc2"/> for the given file path. If
    /// the file is already open in this SW session (case-insensitive
    /// full-path match), returns that handle and reports
    /// <paramref name="weOpenedIt"/> = false (caller MUST NOT
    /// <c>CloseDoc</c> it — that's the user's document). Otherwise
    /// opens silently read-only via <c>OpenDoc6</c> with the requested
    /// config arg, returns the new handle, and reports
    /// <paramref name="weOpenedIt"/> = true (caller SHOULD call
    /// <see cref="CloseIfWeOpened"/> after reading).
    /// </summary>
    /// <param name="app">The SolidWorks application root.</param>
    /// <param name="path">Absolute path to the file. Must exist on disk.</param>
    /// <param name="config">
    /// Config name to load. Pass empty string only when you truly mean
    /// "whichever config was active when this file was last saved" —
    /// see the OpenDoc6 silent-config-arg lesson.
    /// </param>
    /// <param name="weOpenedIt">Output: did we open the file fresh?</param>
    /// <returns>
    /// The doc, or null if SW couldn't open it (returns errors via
    /// <paramref name="errors"/>).
    /// </returns>
    public static IModelDoc2? OpenForReadOrFindOpen(
        ISldWorks app, string path, string config,
        out bool weOpenedIt, out int errors, out int warnings)
    {
        errors = 0;
        warnings = 0;
        weOpenedIt = false;

        // Try the already-open path first. SW exposes a per-title
        // lookup, not a per-path one, so we iterate all open docs and
        // case-insensitive-compare full paths.
        //
        // GetFirstDocument/GetNext return `object` (per the canonical
        // RAG) — we hard-cast inside try/catch rather than using `as
        // IModelDoc2`. The personal_rag lesson
        // `lesson_20260424_as_bodyfolder_cast_unreliable.md` documents
        // that the `as` operator on COM RCWs returned as `object` can
        // silently return null even when the RCW genuinely implements
        // the target interface. Hard-cast goes through QueryInterface
        // and works reliably.
        try
        {
            IModelDoc2? first = null;
            try { first = (IModelDoc2)app.GetFirstDocument(); } catch { first = null; }
            while (first != null)
            {
                try
                {
                    var openPath = first.GetPathName() ?? "";
                    if (string.Equals(openPath, path, StringComparison.OrdinalIgnoreCase))
                    {
                        return first;   // weOpenedIt stays false
                    }
                }
                catch { /* tolerate per-doc property reads */ }

                IModelDoc2? nextDoc = null;
                try { nextDoc = (IModelDoc2)first.GetNext(); } catch { nextDoc = null; }
                first = nextDoc;
            }
        }
        catch { /* fall through to OpenDoc6 */ }

        // Determine doc type from extension. We pass the type through
        // even though OpenDoc6 will accept swDocNONE — being explicit
        // helps SW pick the right engine on the first open and surfaces
        // unsupported extensions as clean errors.
        int docType = DocTypeFromPath(path);
        if (docType == (int)swDocumentTypes_e.swDocNONE) return null;

        // Silent + ReadOnly: don't pop a dialog asking about config
        // mismatches; don't write to the file. ReadOnly is the key
        // primitive that lets us treat this as a query rather than an
        // edit session.
        const int silentReadOnly =
            (int)swOpenDocOptions_e.swOpenDocOptions_Silent |
            (int)swOpenDocOptions_e.swOpenDocOptions_ReadOnly;

        // OpenDoc6's signature returns ModelDoc2 (typed, not object) per
        // the canonical RAG, so the `as`-on-COM-RCW trap doesn't apply
        // here — but we still null-check the result since SW returns
        // null for unreadable files.
        var openedObj = app.OpenDoc6(path, docType, silentReadOnly, config, ref errors, ref warnings);
        IModelDoc2? opened = openedObj as IModelDoc2;
        if (opened is null) return null;
        weOpenedIt = true;
        return opened;
    }

    /// <summary>
    /// Close a doc only if we were the one who opened it. If the user
    /// had it open before our command ran, we leave it alone.
    /// </summary>
    /// <remarks>
    /// Uses <see cref="ISldWorks.CloseDoc"/> by title (NOT path) — the
    /// SW API's only close primitive. Component docs SW loaded as
    /// references remain resident after this call; see the
    /// CloseDoc-leaves-components-resident lesson. We deliberately do
    /// NOT call <c>CloseAllDocuments</c> as a cleanup since that
    /// closes the user's docs too.
    /// </remarks>
    public static void CloseIfWeOpened(ISldWorks app, IModelDoc2 doc, bool weOpenedIt)
    {
        if (!weOpenedIt) return;
        try
        {
            string title = doc.GetTitle() ?? "";
            if (!string.IsNullOrEmpty(title)) app.CloseDoc(title);
        }
        catch { /* nothing useful to do — leave the doc resident */ }
    }

    /// <summary>
    /// Map a file extension to the corresponding <c>swDocumentTypes_e</c>
    /// integer value. Returns <c>swDocNONE</c> for unrecognized
    /// extensions; callers should treat that as a path-validity error.
    /// </summary>
    public static int DocTypeFromPath(string path)
    {
        var ext = Path.GetExtension(path).ToLowerInvariant();
        return ext switch
        {
            ".sldprt" => (int)swDocumentTypes_e.swDocPART,
            ".sldasm" => (int)swDocumentTypes_e.swDocASSEMBLY,
            ".slddrw" => (int)swDocumentTypes_e.swDocDRAWING,
            _         => (int)swDocumentTypes_e.swDocNONE,
        };
    }

    /// <summary>
    /// JSON-escape a string for inclusion in our hand-built output.
    /// We don't use System.Text.Json's serializer at this layer because
    /// the command-level code wants control over field order + the
    /// "empty result vs error" distinction the FR documents; the
    /// command-level layer can then use this for individual string
    /// fields.
    /// </summary>
    public static string JsonString(string? s)
    {
        if (s is null) return "\"\"";
        var sb = new System.Text.StringBuilder(s.Length + 2);
        sb.Append('"');
        foreach (var c in s)
        {
            switch (c)
            {
                case '"':  sb.Append("\\\""); break;
                case '\\': sb.Append("\\\\"); break;
                case '\b': sb.Append("\\b");  break;
                case '\f': sb.Append("\\f");  break;
                case '\n': sb.Append("\\n");  break;
                case '\r': sb.Append("\\r");  break;
                case '\t': sb.Append("\\t");  break;
                default:
                    if (c < ' ') sb.Append($"\\u{(int)c:x4}");
                    else         sb.Append(c);
                    break;
            }
        }
        sb.Append('"');
        return sb.ToString();
    }
}
