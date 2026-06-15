using System.Text;
using ComBridge.Core;
using SolidWorks.Interop.sldworks;

namespace ComBridge.Plugins.SolidWorks;

/// <summary>
/// <c>solidworks active-config</c> — print the active document's
/// active configuration plus the path/title/type as a single JSON
/// object. Designed to be consumed by ScripTree's
/// <c>choices_provider</c> contract or any tool that wants to seed a
/// form with "what's the user currently in?"
/// </summary>
/// <remarks>
/// <para>
/// No file I/O — reads only from the live session. If no document is
/// open, emits the empty-shape JSON
/// (<c>{"path":"","title":"","type":0,"config":""}</c>) with exit 0:
/// "no active doc" is a legitimate read result, not a command error.
/// Callers that want to treat it as failure can check the empty
/// <c>path</c> field.
/// </para>
/// <para>
/// Exit codes: <c>0</c> on success (including the empty-result case).
/// </para>
/// </remarks>
internal sealed class ActiveConfigCommand : IBridgeCommand
{
    public string Name => "active-config";
    public string Usage => "active-config   (prints {path,title,type,config} of the active doc as JSON; empty shape if none)";

    public Task<int> RunAsync(object comRoot, string[] args, TextWriter output)
    {
        var app = (ISldWorks)comRoot;

        string path = "", title = "", config = "";
        int type = 0;

        // ActiveDoc returns object; we hard-cast inside try/catch per
        // the as-on-COM-RCW unreliability lesson.
        IModelDoc2? doc = null;
        try { doc = (IModelDoc2)app.ActiveDoc; } catch { doc = null; }

        if (doc is not null)
        {
            try { path  = doc.GetPathName() ?? ""; } catch { }
            try { title = doc.GetTitle()    ?? ""; } catch { }
            try { type  = doc.GetType(); }            catch { }

            // ConfigurationManager.ActiveConfiguration.Name is the
            // canonical "which config am I in" probe. Each property
            // access can throw on edge cases (no configs at all,
            // detached drawing, etc.) — degrade to empty rather than
            // failing the command.
            try
            {
                var cfgMgr = doc.ConfigurationManager;
                if (cfgMgr is not null)
                {
                    var active = cfgMgr.ActiveConfiguration as IConfiguration;
                    if (active is not null) config = active.Name ?? "";
                }
            }
            catch { }
        }

        // Hand-built JSON. We only emit four primitive fields (string
        // x 3, int x 1) so System.Text.Json's overhead isn't worth it
        // here, and the field order is deterministic for downstream
        // parsing.
        var sb = new StringBuilder();
        sb.Append('{');
        sb.Append("\"path\":")  .Append(SwDocSession.JsonString(path));  sb.Append(',');
        sb.Append("\"title\":") .Append(SwDocSession.JsonString(title)); sb.Append(',');
        sb.Append("\"type\":")  .Append(type);                            sb.Append(',');
        sb.Append("\"config\":").Append(SwDocSession.JsonString(config));
        sb.Append('}');
        output.WriteLine(sb.ToString());
        return Task.FromResult(0);
    }
}
