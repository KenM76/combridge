using System.Diagnostics;
using System.Text;

namespace ComBridge.Plugins.Excel.Mac;

/// <summary>
/// Thin wrapper around <c>osascript</c> (the macOS AppleScript runner).
/// Replaces the COM ROT/oleaut32 attach machinery used by the Windows
/// Excel plugin — on macOS, Excel exposes its scripting surface via
/// AppleEvents/AppleScript, NOT COM, so every "API call" becomes a
/// shell-out to <c>osascript -e '...'</c> with stdout parsed as text.
/// </summary>
/// <remarks>
/// <para>
/// Performance: each call spawns a subprocess (~5-20ms cold). For chatty
/// scripts (iterate 1000 cells one at a time), this is unacceptably slow
/// — batch into one larger AppleScript when possible (the AppleScript can
/// loop internally and emit results in one shot).
/// </para>
/// <para>
/// Quoting: AppleScript source passed via <c>-e</c> gets one round of
/// shell quoting + one round of AppleScript parsing. We use a single
/// <c>-e</c> argument with the AppleScript source as the literal value;
/// .NET's ProcessStartInfo.ArgumentList handles shell quoting for us.
/// Single quotes inside the AppleScript still need careful handling — see
/// <see cref="EscapeForAppleScript"/>.
/// </para>
/// <para>
/// Error model: AppleScript runtime errors come back as stderr text with a
/// non-zero exit code (typically 1). We surface those as
/// <see cref="OsascriptException"/> with the original error preserved.
/// </para>
/// </remarks>
public static class Osascript
{
    /// <summary>
    /// Run an AppleScript expression and return stdout as a trimmed string.
    /// Throws <see cref="OsascriptException"/> on non-zero exit.
    /// </summary>
    public static string Run(string appleScriptSource)
    {
        var psi = new ProcessStartInfo("osascript")
        {
            RedirectStandardOutput = true,
            RedirectStandardError  = true,
            UseShellExecute        = false,
            CreateNoWindow         = true,
        };
        psi.ArgumentList.Add("-e");
        psi.ArgumentList.Add(appleScriptSource);

        using var p = Process.Start(psi)
            ?? throw new InvalidOperationException("Failed to start osascript.");

        var stdout = p.StandardOutput.ReadToEnd();
        var stderr = p.StandardError.ReadToEnd();
        p.WaitForExit();

        if (p.ExitCode != 0)
        {
            throw new OsascriptException(
                $"osascript exited {p.ExitCode}.\nScript:\n{appleScriptSource}\nStderr:\n{stderr.Trim()}");
        }
        return stdout.Trim();
    }

    /// <summary>
    /// Same as <see cref="Run"/> but returns null on error instead of
    /// throwing. Useful for "is X reachable?" probes (e.g. checking
    /// whether Excel is running) where the failure mode is just "no."
    /// </summary>
    public static string? TryRun(string appleScriptSource)
    {
        try { return Run(appleScriptSource); }
        catch (OsascriptException) { return null; }
        catch (InvalidOperationException) { return null; }   // osascript not on PATH
    }

    /// <summary>
    /// Escape a string for safe embedding inside AppleScript double quotes.
    /// AppleScript backslash-escapes <c>"</c> and <c>\</c> inside string
    /// literals. Use as <c>"...\"" + EscapeForAppleScript(s) + "\"..."</c>.
    /// </summary>
    public static string EscapeForAppleScript(string raw)
    {
        var sb = new StringBuilder(raw.Length + 4);
        foreach (var c in raw)
        {
            if (c == '\\' || c == '"') sb.Append('\\');
            sb.Append(c);
        }
        return sb.ToString();
    }

    /// <summary>True if <c>osascript</c> resolves on PATH (i.e. we're on macOS).</summary>
    public static bool IsAvailable()
    {
        // Cheapest possible probe: ask osascript its version. Returns null on
        // any failure including "executable not found."
        return TryRun("return \"ok\"") == "ok";
    }
}

public sealed class OsascriptException : Exception
{
    public OsascriptException(string message) : base(message) { }
}
