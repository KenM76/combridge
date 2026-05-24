namespace ComBridge.Core.Commands;

/// <summary>
/// Built-in <c>run-script</c> command. Every plugin gets this for free —
/// the host registers it before any plugin-specific commands.
/// </summary>
public sealed class RunScriptCommand : IBridgeCommand
{
    private readonly IComBridgePlugin _plugin;
    public RunScriptCommand(IComBridgePlugin plugin) => _plugin = plugin;

    public string Name => "run-script";
    public string Usage => "run-script <scriptFile.csx>   (output file passed as last CLI arg)";

    public async Task<int> RunAsync(object comRoot, string[] args, TextWriter output)
    {
        if (args.Length < 1)
        {
            output.WriteLine($"USAGE: {Usage}");
            return 64;
        }
        var globals = _plugin.CreateGlobals(comRoot);
        return await ScriptHost.RunAsync(_plugin, globals, args[0], output);
    }
}
