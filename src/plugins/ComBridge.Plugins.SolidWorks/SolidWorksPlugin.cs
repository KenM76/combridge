using ComBridge.Core;
using Microsoft.CodeAnalysis;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace ComBridge.Plugins.SolidWorks;

/// <summary>
/// Globals exposed to user .csx scripts. Names mirror the standard SolidWorks
/// API naming convention (swApp, swDoc, swPart, swAssy, swDrawing) so scripts
/// read like the SW macro recorder output and the official sample code.
/// </summary>
public sealed class SwGlobals
{
    public ISldWorks swApp { get; }
    public IModelDoc2? swDoc { get; }
    public IPartDoc? swPart { get; }
    public IAssemblyDoc? swAssy { get; }
    public IDrawingDoc? swDrawing { get; }
    public swDocumentTypes_e swDocType { get; }

    internal SwGlobals(ISldWorks app)
    {
        swApp = app;
        swDoc = app.ActiveDoc as IModelDoc2;
        if (swDoc is not null)
        {
            // IModelDoc2.GetType() returns swDocumentTypes_e as int. Verified
            // against sldworks_methods_v3_llm.rag + swconst_enums.txt:
            //   swDocPART = 1, swDocASSEMBLY = 2, swDocDRAWING = 3.
            swDocType = (swDocumentTypes_e)swDoc.GetType();
            swPart = swDocType == swDocumentTypes_e.swDocPART ? swDoc as IPartDoc : null;
            swAssy = swDocType == swDocumentTypes_e.swDocASSEMBLY ? swDoc as IAssemblyDoc : null;
            swDrawing = swDocType == swDocumentTypes_e.swDocDRAWING ? swDoc as IDrawingDoc : null;
        }
    }
}

public sealed class SolidWorksPlugin : IComBridgePlugin
{
    public string Name => "solidworks";
    public string Description => "SolidWorks (attach to running session via ROT). Globals: swApp, swDoc, swPart, swAssy, swDrawing.";
    public string[] ProgIds => new[] { "SldWorks.Application" };
    public bool AllowCreateNew => false; // SW is heavy — never silently launch it.
    public Type GlobalsType => typeof(SwGlobals);

    // SolidWorks registers each running automation server in the ROT under a
    // unique display name "SolidWorks_PID_<pid>". The bare class moniker
    // "!SldWorks.Application" is also accepted for older SW builds. SW does
    // NOT add itself to the ROT until automation is initialized — typically
    // when the first document loads or an external client first attaches —
    // so an "empty" SW instance with no documents open won't be visible here.
    public IEnumerable<string> RotMonikerPatterns => new[]
    {
        @"^SolidWorks_PID_\d+$",
        @"^!?SldWorks\.Application$",
    };

    public object CreateGlobals(object comRoot) => new SwGlobals((ISldWorks)comRoot);

    public IEnumerable<MetadataReference> ScriptReferences
    {
        get
        {
            // Add every SolidWorks.Interop.* DLL sitting next to this plugin.
            var here = Path.GetDirectoryName(typeof(SolidWorksPlugin).Assembly.Location)!;
            foreach (var dll in Directory.EnumerateFiles(here, "SolidWorks.Interop.*.dll"))
                yield return MetadataReference.CreateFromFile(dll);
        }
    }

    public IEnumerable<string> ScriptUsings => new[]
    {
        "SolidWorks.Interop.sldworks",
        "SolidWorks.Interop.swconst",
        "SolidWorks.Interop.swcommands",
    };

    public IEnumerable<IBridgeCommand> Commands => new IBridgeCommand[]
    {
        new ActiveDocCommand(),
    };

    /// <summary>
    /// Pull the SW main-window HWND via <c>IFrame.GetHWndx64()</c> (verified in
    /// sldworks_methods_v3_llm.rag line 11951), convert to PID via Win32, and
    /// the active-doc title for human identification.
    /// </summary>
    public (int? Pid, string? Title) DescribeInstance(object comRoot)
    {
        try
        {
            var app = (ISldWorks)comRoot;
            int? pid = null;
            try
            {
                var frame = app.Frame() as IFrame;
                if (frame is not null)
                {
                    // GetHWndx64() returns a 64-bit handle on x64; GetHWnd() truncates.
                    var hwnd = (IntPtr)frame.GetHWndx64();
                    pid = SessionPicker.PidFromHwnd(hwnd);
                }
            }
            catch { /* IFrame may be unavailable on a session that's mid-startup */ }

            string? title = null;
            try { title = (app.ActiveDoc as IModelDoc2)?.GetTitle(); } catch { }
            return (pid, title);
        }
        catch
        {
            return (null, null);
        }
    }
}

/// <summary>
/// Quick smoke-test command: prints the title and document type of the active model.
/// Verifies that ROT attach + interop marshalling are working without any user script.
/// </summary>
internal sealed class ActiveDocCommand : IBridgeCommand
{
    public string Name => "active-doc";
    public string Usage => "active-doc   (prints title + type of active document)";

    public Task<int> RunAsync(object comRoot, string[] args, TextWriter output)
    {
        var app = (ISldWorks)comRoot;
        var doc = app.ActiveDoc as IModelDoc2;
        if (doc is null)
        {
            output.WriteLine("(no active document)");
            return Task.FromResult(0);
        }
        var t = (swDocumentTypes_e)doc.GetType();
        output.WriteLine($"Title: {doc.GetTitle()}");
        output.WriteLine($"Path:  {doc.GetPathName()}");
        output.WriteLine($"Type:  {t}");
        return Task.FromResult(0);
    }
}
