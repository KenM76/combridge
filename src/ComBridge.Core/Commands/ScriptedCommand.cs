namespace ComBridge.Core.Commands;

/// <summary>
/// Wraps an auto-discovered <c>.csx</c> file in a plugin's <c>commands/</c>
/// folder as a named <see cref="IBridgeCommand"/>. The command's
/// <see cref="Name"/> is the script filename without the <c>.csx</c>
/// extension; invoking it runs the script against the plugin's globals.
/// </summary>
/// <remarks>
/// Discovered by <see cref="PluginLoader.GetScriptedCommands"/>. Built-in
/// commands (<c>run-script</c>, <c>list-sessions</c>) and the plugin's
/// typed <see cref="IComBridgePlugin.Commands"/> take precedence on name
/// collision — a scripted command can never shadow a plugin-author API.
///
/// <para>This is the implementation of "Shape A" extensibility (per-user
/// command extensions); see <c>LLM/extending.md</c> for the design.</para>
/// </remarks>
public sealed class ScriptedCommand : IBridgeCommand
{
    private readonly IComBridgePlugin _plugin;
    private readonly string _scriptPath;

    /// <summary>The command's CLI name (filename without <c>.csx</c>).</summary>
    public string Name { get; }

    /// <summary>Human-readable usage hint shown by <c>list-commands</c>.</summary>
    public string Usage { get; }

    public ScriptedCommand(IComBridgePlugin plugin, string name, string scriptPath)
    {
        _plugin     = plugin;
        Name        = name;
        _scriptPath = scriptPath;
        Usage       = $"{name}   (.csx: {Path.GetFileName(scriptPath)})";
    }

    public Task<int> RunAsync(object comRoot, string[] args, TextWriter output)
    {
        var globals = _plugin.CreateGlobals(comRoot);
        return ScriptHost.RunAsync(_plugin, globals, _scriptPath, output);
    }
}
