namespace ComBridge.Core.Commands;

/// <summary>
/// Built-in <c>run-script</c> command. Every plugin gets this for free —
/// the host registers it before any plugin-specific commands.
/// </summary>
/// <remarks>
/// <para>
/// CLI shape: <c>combridge &lt;plugin&gt; run-script &lt;script&gt; [args...] &lt;out&gt;</c>.
/// The host strips the trailing output-file positional before calling
/// here, so <c>args[0]</c> is the script path and any further tokens
/// (<c>args[1..]</c>) form the script's <see cref="IScriptContext.ScriptArgs"/>.
/// </para>
/// <para>
/// Stdin behavior: when the calling process redirected stdin to
/// combridge (a pipeline or a here-doc), the full stream is read
/// eagerly at command entry and exposed to the script as
/// <see cref="IScriptContext.Stdin"/>. ScripTree's
/// <c>choices_provider</c> contract (which delivers its input as a
/// JSON object on stdin) is the primary motivator — see
/// <c>FR_runscript_stdin_and_stderr_separation.md</c>.
/// </para>
/// </remarks>
public sealed class RunScriptCommand : IBridgeCommand
{
    private readonly IComBridgePlugin _plugin;
    public RunScriptCommand(IComBridgePlugin plugin) => _plugin = plugin;

    public string Name => "run-script";
    public string Usage => "run-script <scriptFile.{csx|vbs}> [args...]   (output file passed as last CLI arg)";

    public async Task<int> RunAsync(object comRoot, string[] args, TextWriter output)
    {
        if (args.Length < 1)
        {
            output.WriteLine($"USAGE: {Usage}");
            return 64;
        }

        var globals = _plugin.CreateGlobals(comRoot);

        // Populate the host-side I/O channels on globals if the plugin's
        // globals class implements IScriptContext (every shipped plugin
        // does). Plugins that don't implement it just skip silently —
        // their scripts won't see ScriptArgs/Stdin, which matches the
        // pre-v0.9.0 behavior.
        if (globals is IScriptContext ctx)
        {
            // Everything between the script path and the trailing output
            // file goes to the script as ScriptArgs. Empty array if no
            // extra tokens were passed.
            ctx.ScriptArgs = args.Length > 1
                ? args.Skip(1).ToArray()
                : Array.Empty<string>();

            // Read stdin eagerly if it's redirected, so a script can
            // JsonSerializer.Deserialize(Stdin) without stream timing
            // concerns. With a short timeout to avoid hangs when stdin
            // is "redirected" only because a parent shell inherited a
            // non-terminal handle but no producer is actually writing —
            // common when combridge is launched from bash subprocesses
            // or scheduled tasks. Real producers (ScripTree provider
            // pipes, here-docs, file redirects) deliver data within
            // microseconds; 250 ms is generous for them, instant
            // enough to not hang an empty invocation.
            ctx.Stdin = Console.IsInputRedirected
                ? await ReadStdinWithTimeoutAsync(TimeSpan.FromMilliseconds(250))
                : "";
        }

        return await ScriptHost.RunAsync(_plugin, globals, args[0], output);
    }

    /// <summary>
    /// Read the whole of stdin into a string, but give up if no data
    /// arrives within <paramref name="initialWait"/>. After the first
    /// chunk has been read the timer is extended per chunk so a slow
    /// producer can still deliver a large payload; only the INITIAL
    /// wait is bounded.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <see cref="Console.IsInputRedirected"/> returns true whenever
    /// stdin is anything other than a terminal — including when a
    /// parent shell (bash, Task Scheduler, a CI runner) inherits a
    /// non-terminal handle to combridge without writing anything.
    /// A naive <c>Console.In.ReadToEndAsync()</c> blocks forever in
    /// that case because the pipe stays open but empty. This method
    /// uses the underlying stream's
    /// <see cref="System.IO.Stream.ReadAsync(byte[],int,int,System.Threading.CancellationToken)"/>
    /// with a cancellation token so the wait collapses to the timeout
    /// when no producer is on the other end.
    /// </para>
    /// <para>
    /// Real producers (ScripTree provider pipes, Bash here-docs, file
    /// redirects) deliver the first bytes within microseconds — they
    /// never hit the timeout. The 250 ms default is generous; bumping
    /// it further has no upside.
    /// </para>
    /// </remarks>
    private static async Task<string> ReadStdinWithTimeoutAsync(TimeSpan initialWait)
    {
        try
        {
            using var cts = new CancellationTokenSource(initialWait);
            var stream = Console.OpenStandardInput();
            using var ms = new MemoryStream();
            var buf = new byte[8192];
            while (true)
            {
                int read = await stream.ReadAsync(buf.AsMemory(), cts.Token);
                if (read == 0) break;
                ms.Write(buf, 0, read);
                // Producer is actively writing — extend the timer so a
                // multi-chunk payload doesn't get truncated by the
                // initial-wait clock.
                cts.CancelAfter(initialWait);
            }
            return System.Text.Encoding.UTF8.GetString(ms.ToArray());
        }
        catch (OperationCanceledException) { return ""; }
        catch                              { return ""; }
    }
}
