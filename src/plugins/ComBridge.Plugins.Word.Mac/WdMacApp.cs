using System.Globalization;
using ComBridge.Mac.Common;

namespace ComBridge.Plugins.Word.Mac;

/// <summary>
/// Mac Word application wrapper exposed to user scripts as <c>wdApp</c>.
/// Mirrors a subset of the Windows Word plugin's <c>wdApp</c> surface
/// (version, visibility, active document, document enumeration) via
/// AppleScript shell-outs.
/// </summary>
/// <remarks>
/// Application name is <c>"Microsoft Word"</c> (NOT just <c>"Word"</c>).
/// Hardcoded here; would need updating if Microsoft ever renames Word for Mac.
/// </remarks>
public sealed class WdMacApp
{
    private const string AppName = "Microsoft Word";

    /// <summary>Word version string, e.g. "16.85".</summary>
    public string Version =>
        Osascript.Run($"tell application \"{AppName}\" to version");

    /// <summary>Whether Word's UI is visible.</summary>
    public bool Visible
    {
        get
        {
            var raw = Osascript.TryRun(
                $"tell application \"System Events\" to visible of (first process whose name is \"{AppName}\")");
            return raw == "true";
        }
    }

    /// <summary>Number of currently open documents. 0 if Word isn't running.</summary>
    public int DocumentCount
    {
        get
        {
            var raw = Osascript.TryRun($"tell application \"{AppName}\" to count of documents");
            return int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : 0;
        }
    }

    /// <summary>Names of every open document, in Word's enumeration order. Empty if none.</summary>
    public string[] DocumentNames
    {
        get
        {
            var raw = Osascript.TryRun(
                $"tell application \"{AppName}\" to return name of every document as text");
            if (string.IsNullOrEmpty(raw)) return Array.Empty<string>();
            return raw.Split(", ", StringSplitOptions.RemoveEmptyEntries);
        }
    }

    /// <summary>Name of the active document, or null if none open.</summary>
    public string? ActiveDocumentName
    {
        get
        {
            var raw = Osascript.TryRun($"tell application \"{AppName}\" to name of active document");
            return string.IsNullOrEmpty(raw) ? null : raw;
        }
    }

    /// <summary>Full path of the active document, or null if unsaved/no doc.</summary>
    public string? ActiveDocumentPath
    {
        get
        {
            // AppleScript returns a HFS path (with colons); Word also exposes POSIX path on the
            // 'full name' property. Use 'POSIX path of file (...)' for a Unix-style result.
            var raw = Osascript.TryRun(
                $"tell application \"{AppName}\" to POSIX path of (full name of active document)");
            return string.IsNullOrEmpty(raw) ? null : raw;
        }
    }

    /// <summary>
    /// Full text content of the active document. Empty string if no doc / no content.
    /// Word's 'content' on the active document's 'text object' returns the entire body.
    /// </summary>
    public string Content
    {
        get
        {
            var raw = Osascript.TryRun(
                $"tell application \"{AppName}\" to content of text object of active document");
            return raw ?? "";
        }
    }

    /// <summary>
    /// Word count via ComputeStatistics-equivalent. Uses Word's built-in word count.
    /// Returns -1 if unavailable.
    /// </summary>
    public int WordCount
    {
        get
        {
            var raw = Osascript.TryRun(
                $"tell application \"{AppName}\" to count of words of active document");
            return int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : -1;
        }
    }

    /// <summary>Paragraph count.</summary>
    public int ParagraphCount
    {
        get
        {
            var raw = Osascript.TryRun(
                $"tell application \"{AppName}\" to count of paragraphs of active document");
            return int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : -1;
        }
    }
}

/// <summary>
/// Globals exposed to user .csx scripts on macOS. Single field <c>wdApp</c>
/// wraps the Mac AppleScript surface. Parallels the Windows
/// <c>WdGlobals</c> structure (same identifier name, different type).
/// </summary>
public sealed class WdMacGlobals : ComBridge.Core.IScriptContext
{
    public WdMacApp wdApp { get; }

    // Host-injected I/O channels — see ComBridge.Core.IScriptContext.
    public string[] ScriptArgs { get; set; } = Array.Empty<string>();
    public string Stdin { get; set; } = "";

    internal WdMacGlobals() { wdApp = new WdMacApp(); }
}
