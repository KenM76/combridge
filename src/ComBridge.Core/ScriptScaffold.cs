namespace ComBridge.Core;

/// <summary>
/// Helper for plugin <c>new-script &lt;path&gt;</c> commands that scaffold a
/// starter <c>.csx</c> file from a template string.
/// <para>
/// Why this lives in Core: every Office plugin's <c>new-script</c> command
/// does the same mechanical work — parse path + <c>--force</c>, refuse to
/// overwrite existing files unless <c>--force</c>, write the template,
/// report the result. The only thing that varies per plugin is the template
/// text itself. Sharing the dispatch keeps each plugin's <c>new-script</c>
/// command down to a one-line delegate plus its template constant.
/// </para>
/// <para>
/// Design rationale: this is the v0.5.0 replacement for the v0.4.2
/// host-side preamble injection mechanism (<c>ScriptUsingAliases</c>).
/// Both solve the same author-onboarding problem — typing
/// <c>using Xl = global::Microsoft.Office.Interop.Excel;</c> at the top of
/// every Office script — but scaffolding writes the line into the source
/// file the author owns, where every reader (IDE, LLM, app-store auditor,
/// future maintainer) can see it. Preamble injection rewrote the script
/// behind the author's back, breaking the source-is-truth contract. See
/// <c>FR_office_script_interop_alias.md</c> (in the Rejected folder of
/// FeatureRequests) for the full reasoning.
/// </para>
/// </summary>
public static class ScriptScaffold
{
    /// <summary>
    /// Write <paramref name="template"/> to the path supplied in
    /// <paramref name="args"/>. Returns a process-style exit code:
    /// <list type="bullet">
    ///   <item><c>0</c> — file written successfully.</item>
    ///   <item><c>1</c> — destination already exists and <c>--force</c> wasn't passed.</item>
    ///   <item><c>2</c> — directory containing the destination doesn't exist.</item>
    ///   <item><c>64</c> — no path argument supplied (usage error).</item>
    ///   <item><c>5</c> — I/O failure writing the file.</item>
    /// </list>
    /// </summary>
    /// <param name="args">
    /// Raw subcommand argv. Expected form: <c>&lt;path&gt; [--force]</c>.
    /// Unknown flags are tolerated with a warning rather than rejected so
    /// future per-template options (e.g. <c>--name</c>) can be added by
    /// individual plugins without breaking this helper's signature.
    /// </param>
    /// <param name="output">Where to write status messages and warnings.</param>
    /// <param name="template">
    /// The full .csx text to write. Plugins typically embed this as a
    /// <c>const string</c> next to their <c>new-script</c> command class.
    /// </param>
    /// <param name="commandName">CLI name, for usage diagnostics (e.g. "new-script").</param>
    /// <param name="usage">Full usage string the host shows on argument errors.</param>
    public static Task<int> WriteTemplate(
        string[] args,
        TextWriter output,
        string template,
        string commandName,
        string usage)
    {
        string? path = null;
        bool force = false;
        for (int i = 0; i < args.Length; i++)
        {
            string a = args[i];
            switch (a)
            {
                case "--force": force = true; break;
                case "-f":      force = true; break;
                default:
                    if (a.StartsWith("--"))
                    {
                        output.WriteLine($"WARN: unknown flag '{a}' ignored.");
                    }
                    else if (path is null)
                    {
                        path = a;
                    }
                    else
                    {
                        output.WriteLine($"WARN: extra positional argument '{a}' ignored.");
                    }
                    break;
            }
        }

        if (path is null)
        {
            output.WriteLine($"USAGE: {usage}");
            return Task.FromResult(64);
        }

        try
        {
            string full = Path.GetFullPath(path);
            string dir  = Path.GetDirectoryName(full) ?? ".";
            if (!Directory.Exists(dir))
            {
                output.WriteLine($"ERROR: directory does not exist: {dir}");
                output.WriteLine($"       (create it first, or supply a path in an existing folder)");
                return Task.FromResult(2);
            }
            if (File.Exists(full) && !force)
            {
                output.WriteLine($"ERROR: file already exists: {full}");
                output.WriteLine($"       (pass --force to overwrite)");
                return Task.FromResult(1);
            }

            File.WriteAllText(full, template);
            output.WriteLine($"Wrote starter script: {full}");
            output.WriteLine($"  Edit the example body, then run with:");
            output.WriteLine($"    combridge <plugin> run-script {Path.GetFileName(full)} -");
            return Task.FromResult(0);
        }
        catch (IOException ex)
        {
            output.WriteLine($"ERROR writing script: {ex.Message}");
            return Task.FromResult(5);
        }
        catch (UnauthorizedAccessException ex)
        {
            output.WriteLine($"ERROR: access denied: {ex.Message}");
            return Task.FromResult(5);
        }
    }
}
