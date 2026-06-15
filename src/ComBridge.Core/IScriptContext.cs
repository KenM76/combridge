namespace ComBridge.Core;

/// <summary>
/// Host-injected fields available to <c>run-script</c> scripts on top of
/// the plugin-specific globals. Each plugin's globals class implements
/// this interface so the host can set the fields after construction —
/// the plugin owns the COM-binding properties (e.g. <c>swApp</c>,
/// <c>xlBook</c>); the host owns the I/O channels (<c>ScriptArgs</c>,
/// <c>Stdin</c>).
/// </summary>
/// <remarks>
/// <para>
/// Why an interface rather than a base class: each plugin's globals
/// class is already final-state with init-only COM properties set by
/// its constructor. Switching to a base class would force every
/// plugin's globals constructor to chain through, and any future
/// host-added field becomes a versioning concern across the plugin
/// tree. The interface lets the host blindly cast and set without
/// any plugin contract change.
/// </para>
/// <para>
/// Both fields are populated by <see cref="Commands.RunScriptCommand"/>
/// before the script runs. <c>ScriptArgs</c> comes from the CLI tokens
/// between the script path and the output-file positional;
/// <c>Stdin</c> comes from the process's stdin if redirected (empty
/// string otherwise). Scripts that ignore them are unaffected; callers
/// that pass nothing get empty values.
/// </para>
/// <para>
/// See FR <c>FR_runscript_script_args.md</c> (the argv channel) and
/// FR <c>FR_runscript_stdin_and_stderr_separation.md</c> (the stdin
/// channel + stderr separation). Both shipped in v0.9.0.
/// </para>
/// </remarks>
public interface IScriptContext
{
    /// <summary>
    /// CLI tokens passed between the script path and the trailing
    /// output-file positional. Empty array when no tokens were passed.
    /// </summary>
    /// <example>
    /// <c>combridge solidworks run-script audit.csx --mode quick --offline X: -</c>
    /// produces <c>ScriptArgs = ["--mode", "quick", "--offline", "X:"]</c>.
    /// </example>
    string[] ScriptArgs { get; set; }

    /// <summary>
    /// The full stdin the process received, as a single string. Empty
    /// when stdin was not redirected. Read eagerly at script-start so
    /// the script can <c>JsonSerializer.Deserialize</c> it without
    /// worrying about stream timing.
    /// </summary>
    /// <example>
    /// A ScripTree <c>choices_provider</c> hands the .csx its request
    /// as a JSON blob on stdin; the script does
    /// <c>var req = JsonSerializer.Deserialize&lt;Request&gt;(Stdin);</c>
    /// and prints the choices JSON on stdout.
    /// </example>
    string Stdin { get; set; }
}
