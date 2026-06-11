using System.Reflection;
using System.Runtime.InteropServices;
using static ComBridge.Core.ActiveScriptInterop;

namespace ComBridge.Core;

/// <summary>
/// In-process VBScript host for combridge <c>run-script</c>. Runs a
/// <c>.vbs</c> file against the same plugin globals (<c>swApp</c>, etc.)
/// that the Roslyn host exposes to <c>.csx</c> files. Same
/// <c>--session</c> attach, same output capture, same exit-code mapping.
/// </summary>
/// <remarks>
/// <para>
/// Architecture (option B from FR_vbscript_scripting_host.md):
/// <c>CoCreateInstance(CLSID_VBScript)</c> → cast to
/// <see cref="IActiveScript"/> + <see cref="IActiveScriptParse64"/> →
/// install a <see cref="VbScriptSite"/> callback object → enumerate the
/// plugin globals object's public properties and
/// <see cref="IActiveScript.AddNamedItem"/> each one → also add the
/// <c>WScript</c> shim for <c>WScript.Echo</c>/<c>WScript.Quit</c> →
/// <see cref="IActiveScriptParse64.ParseScriptText"/> the script body
/// with <see cref="SCRIPTTEXT_ISVISIBLE"/> so top-level statements run
/// immediately → <see cref="IActiveScript.SetScriptState"/> to
/// <see cref="SCRIPTSTATE_CONNECTED"/> to drive execution.
/// </para>
/// <para>
/// <b>Threading:</b> SolidWorks (and Office) COM is STA. The RCWs in
/// <paramref name="globals"/> are bound to the thread that attached.
/// VBScript engine objects are likewise apartment-affinitized. We do
/// everything on the calling thread — combridge's CLI is already STA per
/// <c>Program.cs</c>'s <c>[STAThread]</c> entry. No marshalling needed.
/// </para>
/// <para>
/// <b>VBScript deprecation:</b> Microsoft announced VBScript's
/// deprecation in 2024 with planned removal from a future Windows
/// release. This host depends on the in-box engine
/// (<c>vbscript.dll</c>). When Microsoft removes it, this command will
/// fail at <see cref="Activator.CreateInstance(Type)"/> with a
/// REGDB_E_CLASSNOTREG error. The <c>.csx</c> path is the long-term-
/// stable option; <c>.vbs</c> exists to run the existing macro corpus
/// while it can still be run.
/// </para>
/// <para>
/// <b>Exit code mapping</b> (parallel to <see cref="ScriptHost"/>):
/// </para>
/// <list type="bullet">
///   <item><c>0</c> — script ran to completion (or called <c>WScript.Quit 0</c>)</item>
///   <item><c>2</c> — script file not found</item>
///   <item><c>3</c> — VBScript syntax/parse error</item>
///   <item><c>4</c> — runtime error (<c>Err.Raise</c>, division by zero,
///         unbound name, COM exception)</item>
///   <item><c>5</c> — host failure (engine CoCreate failed, CLSID not
///         registered, etc.)</item>
///   <item>any other <c>int</c> — value passed to <c>WScript.Quit(N)</c></item>
/// </list>
/// </remarks>
[System.Runtime.Versioning.SupportedOSPlatform("windows")]
internal static class VbScriptEngine
{
    /// <summary>
    /// Execute <paramref name="scriptPath"/> with <paramref name="globals"/>'s
    /// public properties injected as named script items. Returns the
    /// process-style exit code documented on the class.
    /// </summary>
    /// <param name="globals">
    /// The plugin's globals object (e.g. <c>SwGlobals</c>). Every public
    /// instance property becomes a named item visible to the script. A
    /// null property value becomes a script <c>Nothing</c>.
    /// </param>
    /// <param name="scriptPath">Absolute or relative path to a <c>.vbs</c> file.</param>
    /// <param name="output">Sink for <c>WScript.Echo</c> + host diagnostics.</param>
    public static Task<int> RunAsync(
        object globals,
        string scriptPath,
        TextWriter output)
    {
        if (!File.Exists(scriptPath))
        {
            output.WriteLine($"ERROR: script not found: {scriptPath}");
            return Task.FromResult(2);
        }

        // Instantiate the VBScript engine. This is the single point of
        // failure that means "VBScript was removed from this Windows" —
        // we surface it as a host error (exit 5) with a hint.
        object engineRaw;
        try
        {
            var type = Type.GetTypeFromCLSID(CLSID_VBScript);
            if (type is null) throw new InvalidOperationException("Type.GetTypeFromCLSID returned null.");
            engineRaw = Activator.CreateInstance(type)
                ?? throw new InvalidOperationException("Activator.CreateInstance returned null.");
        }
        catch (Exception ex)
        {
            output.WriteLine($"ERROR: failed to create VBScript engine: {ex.Message}");
            output.WriteLine("       (Microsoft deprecated VBScript in 2024; if a recent Windows update");
            output.WriteLine("        removed vbscript.dll, this host will no longer function. Use .csx.)");
            return Task.FromResult(5);
        }

        var engine = (IActiveScript)engineRaw;
        var parser = (IActiveScriptParse64)engineRaw;
        var site   = new VbScriptSite(globals, output);

        try
        {
            int hr;

            hr = parser.InitNew();
            if (hr < 0) throw new COMException("IActiveScriptParse64::InitNew failed", hr);

            hr = engine.SetScriptSite(site);
            if (hr < 0) throw new COMException("IActiveScript::SetScriptSite failed", hr);

            // Register each global as a named script item. The site's
            // GetItemInfo callback hands the engine the actual IDispatch
            // when it asks for a name — we don't pass pointers here, just
            // declarations.
            foreach (var name in site.ItemNames)
            {
                hr = engine.AddNamedItem(name, SCRIPTITEM_ISVISIBLE);
                if (hr < 0) throw new COMException($"AddNamedItem('{name}') failed", hr);
            }

            // Read the whole script. Encoding detection: prefer the BOM,
            // fall back to UTF-8 (default for tool-generated .vbs) — same
            // policy the Roslyn host uses.
            string source = File.ReadAllText(scriptPath, System.Text.Encoding.UTF8);

            // Initialize the engine: required state transition before
            // parsing. Without it, ParseScriptText returns an error.
            hr = engine.SetScriptState(SCRIPTSTATE_INITIALIZED);
            if (hr < 0) throw new COMException("SetScriptState(INITIALIZED) failed", hr);

            // Parse. SCRIPTTEXT_ISVISIBLE puts top-level statements into
            // the global scope so a "WScript.Echo \"hi\"" at file scope
            // actually runs — that's the standard .vbs interpretation.
            // The parser surfaces syntax errors via the EXCEPINFO out
            // param AND via the site's OnScriptError callback (the site
            // captures the error there for richer formatting).
            hr = parser.ParseScriptText(
                pstrCode: source,
                pstrItemName: null,
                punkContext: null,
                pstrDelimiter: null,
                dwSourceContextCookie: 0,
                ulStartingLineNumber: 1,
                dwFlags: SCRIPTTEXT_ISVISIBLE,
                pvarResult: IntPtr.Zero,
                pexcepinfo: out var parseExc);

            if (site.HasParseError || hr < 0)
            {
                // Site already wrote a detailed message via OnScriptError.
                // If it somehow didn't, fall back to the EXCEPINFO struct.
                if (!site.HasParseError && !string.IsNullOrEmpty(parseExc.bstrDescription))
                {
                    output.WriteLine($"PARSE ERROR: {parseExc.bstrDescription}");
                }
                return Task.FromResult(3);
            }

            // Run. Transitioning to CONNECTED executes the parsed code
            // synchronously on the current thread. Runtime errors go
            // through the site's OnScriptError; we read site.HasRuntimeError
            // after the call returns to decide the exit code.
            hr = engine.SetScriptState(SCRIPTSTATE_CONNECTED);
            if (hr < 0 && !site.HasRuntimeError)
            {
                output.WriteLine($"ERROR: SetScriptState(CONNECTED) failed: 0x{hr:X8}");
                return Task.FromResult(5);
            }

            if (site.HasRuntimeError)
                return Task.FromResult(4);
            if (site.HasParseError)
                return Task.FromResult(3);

            // WScript.Quit(N) sets the exit code. If unset, success = 0.
            return Task.FromResult(site.RequestedExitCode ?? 0);
        }
        catch (COMException ex)
        {
            output.WriteLine($"HOST COM EXCEPTION: {ex.Message} (HRESULT 0x{ex.HResult:X8})");
            return Task.FromResult(5);
        }
        catch (Exception ex)
        {
            output.WriteLine($"HOST EXCEPTION: {ex}");
            return Task.FromResult(5);
        }
        finally
        {
            try { engine.Close(); } catch { }
            try { Marshal.FinalReleaseComObject(engineRaw); } catch { }
        }
    }
}

/// <summary>
/// <see cref="IActiveScriptSite"/> implementation: resolves named items
/// from the plugin globals object, captures runtime/parse errors with
/// source context, and hosts the <c>WScript</c> shim (Echo/Quit).
/// </summary>
/// <remarks>
/// <para>
/// The site's responsibility is to answer engine callbacks. The two
/// material ones:
/// </para>
/// <list type="bullet">
///   <item><see cref="GetItemInfo"/> — the engine calls this lazily for
///         each name the script actually references. We hand back the
///         corresponding property value from the globals object (or the
///         <c>WScript</c> shim).</item>
///   <item><see cref="OnScriptError"/> — fired on any parse or runtime
///         error. We format it with line/col + the offending source line
///         and tag whether it was compile-phase or execute-phase so the
///         engine driver can pick exit code 3 vs 4.</item>
/// </list>
/// </remarks>
[System.Runtime.Versioning.SupportedOSPlatform("windows")]
internal sealed class VbScriptSite : IActiveScriptSite
{
    private readonly Dictionary<string, object?> _items;
    private readonly TextWriter _output;
    private readonly WScriptShim _wscript;
    private bool _scriptStartedRunning;

    public bool HasParseError   { get; private set; }
    public bool HasRuntimeError { get; private set; }
    public int? RequestedExitCode => _wscript.ExitCode;
    public IEnumerable<string> ItemNames => _items.Keys;

    public VbScriptSite(object globals, TextWriter output)
    {
        _output  = output;
        _wscript = new WScriptShim(output);
        _items   = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
        {
            // WScript shim is always available; matches cscript convention.
            ["WScript"] = _wscript,
        };

        // Reflect the globals object's public instance properties and
        // expose each one as a named script item. Three filters:
        //   - Must be readable, non-indexer (the obvious shape filter).
        //   - Skip value types (boxed primitives, enums). The
        //     IActiveScript site contract returns globals as IUnknown for
        //     IDispatch wrapping; boxed primitives can't be returned that
        //     way. Live-test caught this on SwGlobals.swDocType (the
        //     swDocumentTypes_e enum). Scripts that need the value can
        //     read it off the typed wrapper (e.g. swDoc.GetType()).
        //   - Skip null reference-type values. EMPIRICAL FINDING:
        //     registering a name via AddNamedItem and then returning
        //     S_OK + ppiunkItem=null from GetItemInfo crashes the
        //     VBScript engine with access violation 0xC0000005 inside
        //     SetScriptState(SCRIPTSTATE_CONNECTED) — verified live
        //     2026-06-11 against vbscript.dll on Win11 26200. The docs'
        //     "If the item cannot be located, this parameter is set to
        //     NULL" wording suggested null might mean Nothing, but it
        //     doesn't — the engine simply isn't defensive against null
        //     dispatch and dereferences it. So null-reference globals
        //     genuinely have to be skipped. The cost is real: scripts
        //     can't use `Option Explicit` (the name is undeclared) and
        //     must use `If Not IsObject(swDrawing)` rather than the
        //     idiomatic `If swDrawing Is Nothing`. We accept that cost.
        var skippedValueType = new List<string>();
        var skippedNull      = new List<string>();
        foreach (var prop in globals.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance))
        {
            if (!prop.CanRead) continue;
            if (prop.GetIndexParameters().Length > 0) continue;

            // Underlying type check covers Nullable<T> wrappers too.
            var underlying = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;
            if (underlying.IsValueType)
            {
                skippedValueType.Add(prop.Name);
                continue;
            }

            object? value;
            try { value = prop.GetValue(globals); }
            catch { value = null; }
            if (value is null)
            {
                skippedNull.Add(prop.Name);
                continue;
            }
            _items[prop.Name] = value;
        }

        // Surface the skip lists so authors aren't mystified when a name
        // they expect comes back as "not defined" at runtime. The host
        // writes these on first run BEFORE the script executes so they
        // appear above any script output in the log.
        if (skippedValueType.Count > 0)
            output.WriteLine($"# vbscript host: value-type globals not injected (read via typed wrapper): {string.Join(", ", skippedValueType)}");
        if (skippedNull.Count > 0)
            output.WriteLine($"# vbscript host: null globals not injected (no active doc/etc. — guard with `If Not IsObject(<name>)`, not `If <name> Is Nothing`): {string.Join(", ", skippedNull)}");
    }

    int IActiveScriptSite.GetLCID(out uint plcid)
    {
        plcid = 0;          // 0 = "use system default"
        return 0;
    }

    int IActiveScriptSite.GetItemInfo(
        string pstrName, uint dwReturnMask,
        out object? ppiunkItem, IntPtr ppti)
    {
        ppiunkItem = null;
        if ((dwReturnMask & SCRIPTINFO_IUNKNOWN) == 0)
            return unchecked((int)0x80070057);  // E_INVALIDARG — we only do IUNKNOWN

        if (!_items.TryGetValue(pstrName, out var value) || value is null)
        {
            // Engine asked for a name we didn't register, or the property
            // was null. TYPE_E_ELEMENTNOTFOUND is the documented response.
            // We MUST NOT return S_OK with ppiunkItem=null here — that
            // crashes the VBScript engine with 0xC0000005 in
            // SetScriptState(CONNECTED). Verified live 2026-06-11.
            return unchecked((int)0x8002802B);
        }

        ppiunkItem = value;
        return 0;
    }

    int IActiveScriptSite.GetDocVersionString(out string? pbstrVersion)
    {
        pbstrVersion = null;
        return unchecked((int)0x80004001);  // E_NOTIMPL
    }

    int IActiveScriptSite.OnScriptTerminate(ref object pvarResult, ref EXCEPINFO pexcepinfo) => 0;

    int IActiveScriptSite.OnStateChange(uint state) => 0;

    int IActiveScriptSite.OnScriptError(IActiveScriptError pscripterror)
    {
        try
        {
            pscripterror.GetExceptionInfo(out var exc);
            pscripterror.GetSourcePosition(out _, out var line, out var col);
            pscripterror.GetSourceLineText(out var lineText);

            // Distinguishing compile vs runtime: if the engine has not yet
            // reached the CONNECTED state when the error fires, it's a
            // parse-time error (exit 3). Otherwise it's runtime (exit 4).
            bool isRuntime = _scriptStartedRunning;
            if (isRuntime) HasRuntimeError = true;
            else           HasParseError   = true;

            string phase = isRuntime ? "RUNTIME ERROR" : "PARSE ERROR";
            _output.WriteLine($"{phase} at line {line}, col {col}: {exc.bstrDescription?.Trim()}");
            if (!string.IsNullOrEmpty(lineText))
                _output.WriteLine($"  > {lineText.TrimEnd()}");
        }
        catch (Exception ex)
        {
            _output.WriteLine($"OnScriptError handler threw: {ex.Message}");
            HasRuntimeError = true;
        }
        return 0;
    }

    /// <summary>
    /// Fires when ANY code unit begins executing — global scope, a sub,
    /// a function. We use this (rather than OnStateChange(CONNECTED),
    /// which only fires when the script COMPLETES) to track "execution
    /// has started" for parse-vs-runtime phase classification in
    /// <see cref="OnScriptError"/>. Caught by the live test: a runtime
    /// division-by-zero was mis-labeled "PARSE ERROR" because the
    /// CONNECTED state-change fires too late to drive phase detection.
    /// </summary>
    int IActiveScriptSite.OnEnterScript()
    {
        _scriptStartedRunning = true;
        return 0;
    }

    int IActiveScriptSite.OnLeaveScript() => 0;
}

/// <summary>
/// Minimal <c>WScript</c> shim exposed to user scripts as a global. Mirrors
/// just enough of <c>cscript.exe</c>'s WScript object to make
/// <c>WScript.Echo "..."</c> and <c>WScript.Quit N</c> work — the two
/// idioms that appear in essentially every standalone-VBScript SW macro.
/// </summary>
/// <remarks>
/// Deliberately omitted vs <c>cscript</c>: <c>WScript.Arguments</c>,
/// <c>WScript.CreateObject</c>, <c>WScript.GetObject</c>,
/// <c>WScript.Sleep</c>, <c>WScript.StdOut</c>. The FR proposed
/// <c>WScript.Arguments</c>; we deferred it to keep v0.8.0 scoped to
/// the FifiRiri acceptance test and to avoid pre-building cscript-compat
/// surface that no actual script needs. <c>ScriptArgs</c> (the channel
/// already shared with .csx) is the supported way to pass argv to
/// scripts.
/// </remarks>
[ComVisible(true)]
[ClassInterface(ClassInterfaceType.AutoDispatch)]
[System.Runtime.Versioning.SupportedOSPlatform("windows")]
public class WScriptShim
{
    private readonly TextWriter _output;
    internal int? ExitCode { get; private set; }

    internal WScriptShim(TextWriter output) { _output = output; }

    /// <summary>
    /// <c>WScript.Echo</c> — writes one line to the combridge output
    /// sink. Accepts any object; uses <c>ToString()</c>. Matches cscript
    /// behavior of newline-terminating each Echo call.
    /// </summary>
    public void Echo(object? message)
    {
        _output.WriteLine(message?.ToString() ?? "");
    }

    /// <summary>
    /// <c>WScript.Quit N</c> — sets the requested exit code. The script
    /// continues executing to end-of-scope; combridge returns the
    /// requested code after the engine completes (not interrupted —
    /// real cscript truly halts, but most macros call Quit at end-of-flow
    /// anyway, so this is correct in practice).
    /// </summary>
    public void Quit(int code = 0)
    {
        ExitCode = code;
    }
}
