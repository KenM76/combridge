using ComBridge.Core;
using ComBridge.Core.Commands;

namespace ComBridge.Cli;

/// <summary>
/// Entry point for <c>combridge.exe</c> — the host process that loads
/// plugins from <c>./plugins/&lt;Name&gt;/</c> and dispatches CLI arguments
/// to the matching plugin's <see cref="IBridgeCommand"/> or built-in
/// commands (<c>list-plugins</c>, <c>list-sessions</c>, <c>run-script</c>).
/// </summary>
/// <remarks>
/// Full CLI grammar, exit codes, and selector behavior lives in
/// <c>LLM/cli.md</c>; the human-facing usage is in <c>README.md</c>.
/// </remarks>
internal static class Program
{
    /// <summary>
    /// Main entry. Parses CLI args, loads plugins, dispatches to the requested
    /// command, returns the command's exit code.
    /// </summary>
    /// <param name="args">
    /// Command-line arguments. Shape:
    /// <list type="bullet">
    ///   <item><c>combridge list-plugins</c></item>
    ///   <item><c>combridge &lt;plugin&gt; list-commands</c></item>
    ///   <item><c>combridge &lt;plugin&gt; list-sessions &lt;outputFile&gt;</c></item>
    ///   <item><c>combridge &lt;plugin&gt; [--session &lt;sel&gt;] [--no-create] &lt;command&gt; [cmd-args...] &lt;outputFile&gt;</c></item>
    /// </list>
    /// The last positional argument is always the output file (<c>-</c> = stdout).
    /// <c>--session</c> selector forms: <c>N</c> (1-based index), <c>pid:NNNN</c>,
    /// or any other string (case-insensitive title substring). <c>--no-create</c>
    /// forbids spawning a new app instance even if the plugin's
    /// <see cref="IComBridgePlugin.AllowCreateNew"/> is <c>true</c>.
    /// </param>
    /// <returns>
    /// Process exit code. See <c>LLM/cli.md</c> § "Exit codes" for the full table:
    /// 0 = success, 1 = unhandled exception, 2 = script not found,
    /// 3 = script compile error, 4 = script runtime exception,
    /// 5 = host exception, 6 = could not connect / no session matched,
    /// 64 = usage error.
    /// </returns>
    private static async Task<int> Main(string[] args)
    {
        try
        {
            if (args.Length == 0 || args[0] is "--help" or "-h" or "help")
            {
                PrintTopLevelHelp();
                return 0;
            }

            var noCreate = false;
            string? sessionSelector = null;
            var filtered = new List<string>();
            for (int i = 0; i < args.Length; i++)
            {
                if (args[i] == "--no-create") noCreate = true;
                else if (args[i] == "--session" && i + 1 < args.Length) sessionSelector = args[++i];
                else if (args[i].StartsWith("--session=")) sessionSelector = args[i]["--session=".Length..];
                else filtered.Add(args[i]);
            }
            args = filtered.ToArray();

            if (args[0] == "list-plugins")
            {
                foreach (var p in PluginLoader.LoadAll())
                    Console.WriteLine($"  {p.Name,-14} {p.Description}");
                return 0;
            }

            if (args.Length < 2)
            {
                PrintTopLevelHelp();
                return 64;
            }

            var pluginName = args[0];
            var commandName = args[1];
            var rest = args.Skip(2).ToArray();

            var plugins = PluginLoader.LoadAll();
            var plugin = plugins.FirstOrDefault(p =>
                string.Equals(p.Name, pluginName, StringComparison.OrdinalIgnoreCase));
            if (plugin is null)
            {
                Console.Error.WriteLine($"ERROR: no plugin named '{pluginName}'. Available:");
                foreach (var p in plugins) Console.Error.WriteLine($"  {p.Name}");
                return 64;
            }

            if (commandName == "list-commands")
            {
                Console.WriteLine($"  run-script    (built-in)  run-script <scriptFile.csx>");
                Console.WriteLine($"  list-sessions (built-in)  list running instances of this plugin's app");
                foreach (var c in plugin.Commands)
                    Console.WriteLine($"  {c.Name,-13} {c.Usage}");
                return 0;
            }

            if (commandName == "list-sessions")
            {
                var sessions = SessionPicker.Enumerate(plugin);
                if (sessions.Count == 0)
                {
                    Console.WriteLine($"(no running {plugin.Name} sessions in the ROT)");
                    return 0;
                }
                Console.WriteLine($"Running {plugin.Name} sessions:");
                foreach (var (_, info) in sessions)
                    Console.WriteLine($"  #{info.Index}  {info.Description}");
                Console.WriteLine();
                Console.WriteLine("Select with --session N (1=MRU) | --session pid:NNNN | --session <title> | --session last");
                Console.WriteLine("Default with no --session = MRU (most-recently-focused window).");
                return 0;
            }

            // Last arg is the output file (or stdout if "-").
            if (rest.Length < 1)
            {
                Console.Error.WriteLine("ERROR: missing output file (last arg). Use '-' for stdout.");
                return 64;
            }
            var outputFile = rest[^1];
            var cmdArgs = rest[..^1];

            // Built-in run-script + plugin commands.
            IBridgeCommand? command = commandName == "run-script"
                ? new RunScriptCommand(plugin)
                : plugin.Commands.FirstOrDefault(c =>
                    string.Equals(c.Name, commandName, StringComparison.OrdinalIgnoreCase));

            if (command is null)
            {
                Console.Error.WriteLine($"ERROR: plugin '{plugin.Name}' has no command '{commandName}'.");
                return 64;
            }

            // Attach (or create) the COM root before invoking the command.
            // Both the selector and the no-selector paths route through
            // SessionPicker.Enumerate so the result is MRU-sorted (default is
            // the most-recently-focused window). Activator.CreateInstance is
            // the fallback only when no live session is discoverable AND the
            // plugin opts into AllowCreateNew (and the user didn't pass --no-create).
            object comRoot;
            try
            {
                var sessions = SessionPicker.Enumerate(plugin);

                if (sessionSelector is not null)
                {
                    var picked = SessionPicker.Resolve(sessions, sessionSelector);
                    if (picked is null)
                    {
                        Console.Error.WriteLine(
                            $"ERROR: --session '{sessionSelector}' matched no running {plugin.Name} instance.");
                        if (sessions.Count > 0)
                        {
                            Console.Error.WriteLine("Available:");
                            foreach (var (_, info) in sessions)
                                Console.Error.WriteLine($"  #{info.Index}  {info.Description}");
                        }
                        return 6;
                    }
                    comRoot = picked.Value.Root;
                }
                else if (sessions.Count > 0)
                {
                    // No --session: pick the MRU (most-recently-focused) session.
                    // SessionPicker.Enumerate sorts MRU-first, so sessions[0] is correct.
                    comRoot = sessions[0].Root;
                }
                else
                {
                    // No live sessions at all → fall through to RotHelper.AttachOrCreate's
                    // create-new-instance path (Activator + ProgID lookup). It does its
                    // own attach attempt first (cheap, no-op when sessions is empty) and
                    // then creates the instance if allowed.
                    comRoot = RotHelper.AttachOrCreate(
                        plugin.RotMonikerPatterns,
                        plugin.ProgIds,
                        transformer: plugin.TryExtractRoot,
                        createIfMissing: plugin.AllowCreateNew && !noCreate);
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"ERROR: could not connect to {plugin.Name}: {ex.Message}");
                return 6;
            }

            using TextWriter writer = outputFile == "-"
                ? Console.Out
                : new StreamWriter(outputFile);
            try
            {
                return await command.RunAsync(comRoot, cmdArgs, writer);
            }
            finally
            {
                writer.Flush();
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("FATAL: " + ex);
            return 1;
        }
    }

    private static void PrintTopLevelHelp()
    {
        Console.WriteLine("combridge - generic COM-automation host with plugins");
        Console.WriteLine();
        Console.WriteLine("Usage:");
        Console.WriteLine("  combridge list-plugins");
        Console.WriteLine("  combridge <plugin> list-commands");
        Console.WriteLine("  combridge <plugin> <command> [args...] <outputFile>");
        Console.WriteLine();
        Console.WriteLine("Flags:");
        Console.WriteLine("  --no-create        Don't launch a new app instance if none is running.");
        Console.WriteLine("  --session <sel>    Pick a specific running instance:");
        Console.WriteLine("                       N           1-based index from `list-sessions`");
        Console.WriteLine("                       pid:NNNN    match by Win32 PID");
        Console.WriteLine("                       <substr>    case-insensitive title substring");
        Console.WriteLine();
        Console.WriteLine("Built-in commands (every plugin):");
        Console.WriteLine("  run-script <scriptFile.csx>");
        Console.WriteLine("  list-sessions");
    }
}
