using System.Text;
using ComBridge.Core;
using SolidWorks.Interop.sldworks;

namespace ComBridge.Plugins.SolidWorks;

/// <summary>
/// <c>solidworks list-configs [path]</c> — list a file's configurations
/// as JSON: <c>{"path","active","configs":[{"name","is_derived"}...]}</c>.
/// </summary>
/// <remarks>
/// <para>
/// Behavior matrix:
/// </para>
/// <list type="bullet">
///   <item><b>Path given, file already open</b>: reads the live doc's
///         configuration list without disturbing it.</item>
///   <item><b>Path given, file NOT open</b>: silently opens read-only,
///         reads, then <c>CloseDoc</c>s. Cost ~3 sec per file. The
///         OpenDoc6 silent-config-arg trap is sidestepped here by
///         passing <c>""</c> for the config — list-configs only reads
///         the configuration NAMES (which are identical regardless of
///         which config is active), so we genuinely don't care which
///         config loads. Per
///         <c>lesson_20260512_opendoc6_config_arg_silent_bug.md</c>,
///         this is the ONE legitimate "" use case.</item>
///   <item><b>Path omitted</b>: uses the active doc (per
///         <c>lesson_20260427_active_doc_fallback_pattern.md</c>). If
///         there's no active doc, emits an empty result with exit 1.</item>
/// </list>
/// <para>
/// Component-residual caveat (per
/// <c>lesson_20260608_closedoc_leaves_components_resident.md</c>):
/// <c>ISldWorks.CloseDoc</c> only closes the named top-level doc.
/// Components that SW loaded as references for that doc remain
/// resident in the session. We deliberately do NOT call
/// <c>CloseAllDocuments(true)</c> as cleanup because it would close
/// the USER'S work with unsaved changes destroyed. Power users who
/// invoke this many times against different assemblies in one
/// session may see component-doc growth; that's the tradeoff for
/// keeping their work safe.
/// </para>
/// <para>
/// Exit codes: <c>0</c> success; <c>1</c> path doesn't exist, has
/// unrecognized extension, or no path AND no active doc.
/// </para>
/// </remarks>
internal sealed class ListConfigsCommand : IBridgeCommand
{
    public string Name => "list-configs";
    public string Usage => "list-configs [<path>]   (lists configs as JSON; omit path to use active doc)";

    public Task<int> RunAsync(object comRoot, string[] args, TextWriter output)
    {
        var app = (ISldWorks)comRoot;
        string? path = args.Length >= 1 ? args[0] : null;

        // Three acquisition paths converge on "we have an IModelDoc2".
        IModelDoc2? doc = null;
        bool weOpenedIt = false;
        string resolvedPath;

        if (string.IsNullOrEmpty(path))
        {
            // Active-doc fallback.
            try { doc = (IModelDoc2)app.ActiveDoc; } catch { doc = null; }
            if (doc is null)
            {
                output.WriteLine("ERROR: no path given and no active document.");
                return Task.FromResult(1);
            }
            try { resolvedPath = doc.GetPathName() ?? ""; }
            catch { resolvedPath = ""; }
        }
        else
        {
            if (!File.Exists(path))
            {
                output.WriteLine($"ERROR: file not found: {path}");
                return Task.FromResult(1);
            }
            if (SwDocSession.DocTypeFromPath(path) == 0)
            {
                output.WriteLine($"ERROR: unrecognized SOLIDWORKS file extension: {Path.GetExtension(path)}");
                return Task.FromResult(1);
            }

            // OpenForReadOrFindOpen reuses the live doc if one matches
            // the path, otherwise silently opens read-only. We pass ""
            // for config because list-configs doesn't care which config
            // is active — the NAMES are config-independent.
            doc = SwDocSession.OpenForReadOrFindOpen(
                app, path, config: "", out weOpenedIt, out int errs, out int warns);
            if (doc is null)
            {
                output.WriteLine($"ERROR: OpenDoc6 returned null (errors=0x{errs:X}, warnings=0x{warns:X}).");
                return Task.FromResult(1);
            }
            resolvedPath = path;
        }

        try
        {
            // GetConfigurationNames returns an object that wraps a
            // BSTR safe-array; in C# interop it materializes as
            // string[]. We tolerate the null/type-mismatch cases by
            // returning an empty list rather than crashing.
            string[] names = Array.Empty<string>();
            try
            {
                var raw = doc.GetConfigurationNames();
                if (raw is string[] s) names = s;
            }
            catch { }

            // Active config: same read as in active-config command.
            string activeName = "";
            try
            {
                var cfgMgr = doc.ConfigurationManager;
                if (cfgMgr is not null)
                {
                    var active = cfgMgr.ActiveConfiguration as IConfiguration;
                    if (active is not null) activeName = active.Name ?? "";
                }
            }
            catch { }

            // Build the configs[] array. For each name, look up the
            // typed Configuration and read IsDerived. Errors collapse
            // to is_derived=false rather than excluding the row, so
            // the caller's array index lines up with the original name
            // ordering.
            var sb = new StringBuilder();
            sb.Append('{');
            sb.Append("\"path\":")  .Append(SwDocSession.JsonString(resolvedPath)); sb.Append(',');
            sb.Append("\"active\":").Append(SwDocSession.JsonString(activeName));   sb.Append(',');
            sb.Append("\"configs\":[");
            for (int i = 0; i < names.Length; i++)
            {
                if (i > 0) sb.Append(',');
                string n = names[i];
                bool isDerived = false;
                try
                {
                    // GetConfigurationByName returns object (per RAG) —
                    // hard-cast inside try/catch, not `as`, per the
                    // as-on-COM-RCW unreliability lesson.
                    var cfgObj = doc.GetConfigurationByName(n);
                    IConfiguration? cfg = null;
                    try { cfg = (IConfiguration)cfgObj; } catch { cfg = null; }
                    if (cfg is not null) isDerived = cfg.IsDerived();
                }
                catch { }
                sb.Append('{');
                sb.Append("\"name\":")      .Append(SwDocSession.JsonString(n)); sb.Append(',');
                sb.Append("\"is_derived\":").Append(isDerived ? "true" : "false");
                sb.Append('}');
            }
            sb.Append("]}");
            output.WriteLine(sb.ToString());
            return Task.FromResult(0);
        }
        finally
        {
            SwDocSession.CloseIfWeOpened(app, doc, weOpenedIt);
        }
    }
}
