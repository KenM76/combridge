using System.Globalization;
using ComBridge.Mac.Common;

namespace ComBridge.Plugins.PowerPoint.Mac;

/// <summary>
/// Mac PowerPoint application wrapper exposed to user scripts as
/// <c>pptApp</c>. Mirrors a subset of the Windows PowerPoint plugin's
/// surface (version, active presentation, slide enumeration) via
/// AppleScript shell-outs.
/// </summary>
public sealed class PptMacApp
{
    private const string AppName = "Microsoft PowerPoint";

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

    public int PresentationCount
    {
        get
        {
            var raw = Osascript.TryRun($"tell application \"{AppName}\" to count of presentations");
            return int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : 0;
        }
    }

    public string? ActivePresentationName
    {
        get
        {
            var raw = Osascript.TryRun($"tell application \"{AppName}\" to name of active presentation");
            return string.IsNullOrEmpty(raw) ? null : raw;
        }
    }

    public string? ActivePresentationPath
    {
        get
        {
            var raw = Osascript.TryRun(
                $"tell application \"{AppName}\" to POSIX path of (full name of active presentation)");
            return string.IsNullOrEmpty(raw) ? null : raw;
        }
    }

    public int SlideCount
    {
        get
        {
            var raw = Osascript.TryRun(
                $"tell application \"{AppName}\" to count of slides of active presentation");
            return int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : 0;
        }
    }

    /// <summary>
    /// Titles of each slide (best-effort — slides without title placeholders
    /// yield the empty string at that index). One osascript call.
    /// </summary>
    public string[] SlideTitles
    {
        get
        {
            // Loop in AppleScript to avoid N round-trips. Build a tab-joined string.
            var script = $@"
tell application ""{AppName}""
    set theSlides to slides of active presentation
    set AppleScript's text item delimiters to tab
    set out to {{}}
    repeat with s in theSlides
        try
            set t to text frame of placeholder 1 of s
            set end of out to (content of text range of t) as string
        on error
            set end of out to """"
        end try
    end repeat
    return out as text
end tell".Trim();
            var raw = Osascript.TryRun(script);
            if (string.IsNullOrEmpty(raw)) return Array.Empty<string>();
            return raw.Split('\t');
        }
    }

    /// <summary>1-based index of the currently-selected slide in the active window.</summary>
    public int? ActiveSlideIndex
    {
        get
        {
            var raw = Osascript.TryRun(
                $"tell application \"{AppName}\" to slide index of slide range of selection of document window 1");
            return int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : null;
        }
    }
}

public sealed class PptMacGlobals
{
    public PptMacApp pptApp { get; }
    internal PptMacGlobals() { pptApp = new PptMacApp(); }
}
