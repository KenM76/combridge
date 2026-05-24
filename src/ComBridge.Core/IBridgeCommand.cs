namespace ComBridge.Core;

/// <summary>
/// A plugin-supplied command callable as <c>combridge &lt;plugin&gt; &lt;command&gt; ...</c>.
/// </summary>
public interface IBridgeCommand
{
    string Name { get; }
    string Usage { get; }

    /// <summary>
    /// Execute the command. <paramref name="comRoot"/> is the connected COM root
    /// object (already attached/created by the host before this call).
    /// </summary>
    /// <returns>Process exit code (0 = success).</returns>
    Task<int> RunAsync(object comRoot, string[] args, TextWriter output);
}
