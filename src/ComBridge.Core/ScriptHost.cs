using System.Text.RegularExpressions;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Scripting;
using Microsoft.CodeAnalysis.Scripting;
using Microsoft.CodeAnalysis.Scripting.Hosting;

namespace ComBridge.Core;

/// <summary>
/// Compiles and runs a user .csx file against a plugin's globals object.
/// The script can <c>Console.WriteLine</c> freely; output is captured to the writer.
/// </summary>
public static class ScriptHost
{
    /// <summary>
    /// Compile and execute a user .csx script against <paramref name="globals"/>
    /// (typed as <see cref="IComBridgePlugin.GlobalsType"/>). Script output via
    /// <see cref="Console.Out"/> and <see cref="Console.Error"/> is redirected to
    /// <paramref name="output"/> for the duration of the run.
    /// </summary>
    /// <param name="plugin">
    /// The plugin whose <see cref="IComBridgePlugin.ScriptReferences"/>,
    /// <see cref="IComBridgePlugin.ScriptUsings"/>, and
    /// <see cref="IComBridgePlugin.GlobalsType"/> shape the compile environment.
    /// The plugin's own assembly and the globals' assembly are registered with
    /// Roslyn's <c>InteractiveAssemblyLoader</c> so user scripts cast cleanly
    /// across the plugin's <c>AssemblyLoadContext</c> boundary.
    /// </param>
    /// <param name="globals">
    /// Instance the script binds against. Must be assignable to
    /// <paramref name="plugin"/>'s <see cref="IComBridgePlugin.GlobalsType"/>.
    /// </param>
    /// <param name="scriptPath">Absolute or relative path to the .csx file.</param>
    /// <param name="output">Receives both stdout-equivalent and stderr-equivalent script output, plus host diagnostics.</param>
    /// <returns>
    /// Process exit code:
    /// <list type="bullet">
    ///   <item><c>0</c> — script completed successfully (or returned 0 explicitly).</item>
    ///   <item><c>2</c> — script file not found at <paramref name="scriptPath"/>.</item>
    ///   <item><c>3</c> — Roslyn compilation produced one or more errors.</item>
    ///   <item><c>4</c> — script ran but threw an exception (<c>state.Exception</c>).</item>
    ///   <item><c>5</c> — host exception during script execution (e.g. ALC / loader failure).</item>
    ///   <item>any other <c>int</c> — value returned by the script's top-level <c>return</c>.</item>
    /// </list>
    /// </returns>
    /// <remarks>
    /// The .csx file is opened as a Stream (not read as a string) so Roslyn
    /// can detect its encoding and emit PDB debug info — passing a string
    /// triggers <c>error CS8055: Cannot emit debug information for a source text
    /// without encoding</c> when the file lacks a BOM (common for tool-generated
    /// .csx). <c>Microsoft.CSharp</c>, <c>System.Runtime</c> (DynamicAttribute),
    /// and <c>System.Linq.Expressions</c> (CallSite) are included in the default
    /// script reference set so user scripts can use <c>dynamic</c> for late-bound
    /// IDispatch calls on Office COM objects without extra setup.
    /// </remarks>
    public static async Task<int> RunAsync(
        IComBridgePlugin plugin,
        object globals,
        string scriptPath,
        TextWriter output)
    {
        if (!File.Exists(scriptPath))
        {
            output.WriteLine($"ERROR: script not found: {scriptPath}");
            return 2;
        }

        // Dispatch by extension. New engines are added here — each
        // engine consumes the same plugin globals object via its own
        // injection mechanism (Roslyn binds typed; VBScript injects via
        // AddNamedItem + reflection). See FR_vbscript_scripting_host.md
        // for the architectural reasoning.
        string ext = Path.GetExtension(scriptPath).ToLowerInvariant();
        switch (ext)
        {
            case ".csx":
                break;  // fall through to Roslyn path below
            case ".vbs":
                if (!OperatingSystem.IsWindows())
                {
                    output.WriteLine("ERROR: .vbs scripts require Windows (in-box VBScript engine).");
                    return 5;
                }
                return await VbScriptEngine.RunAsync(globals, scriptPath, output);
            case ".vba":
            case ".bas":
            case ".swp":
                output.WriteLine($"ERROR: {ext} files are VBA, not VBScript — combridge can't host them.");
                output.WriteLine($"       VBA uses UserForms, Type definitions, Public Const etc. that the");
                output.WriteLine($"       VBScript engine won't parse. Convert to .vbs (delete attach line,");
                output.WriteLine($"       remove UserForm dependencies), or rewrite as .csx.");
                return 5;
            default:
                output.WriteLine($"ERROR: unknown script extension '{ext}'. Supported: .csx, .vbs");
                return 5;
        }

        // (Script file is opened as a Stream below, so Roslyn picks up the
        // encoding from a BOM or falls back to UTF-8. Passing a string would
        // drop encoding info and trigger:
        //   error CS8055: Cannot emit debug information for a source text
        //   without encoding.)

        var refs = new List<MetadataReference>(plugin.ScriptReferences);
        // Always include core BCL + the plugin's own assembly so its globals type resolves.
        refs.Add(MetadataReference.CreateFromFile(typeof(object).Assembly.Location));
        refs.Add(MetadataReference.CreateFromFile(typeof(System.Linq.Enumerable).Assembly.Location));
        refs.Add(MetadataReference.CreateFromFile(typeof(System.Console).Assembly.Location));
        refs.Add(MetadataReference.CreateFromFile(plugin.GetType().Assembly.Location));
        refs.Add(MetadataReference.CreateFromFile(plugin.GlobalsType.Assembly.Location));
        // For `dynamic` support in user scripts (late-bound IDispatch calls on
        // Office objects, etc.). The C# compiler needs three things, none in
        // ScriptOptions.Default:
        //   - Microsoft.CSharp.dll        : the RuntimeBinder
        //   - DynamicAttribute            : lives in System.Runtime / netstandard
        //   - CallSite + CallSiteBinder   : in System.Core / System.Linq.Expressions
        // Without all three, `dynamic` triggers CS0656 / CS1980 / CS0518.
        refs.Add(MetadataReference.CreateFromFile(typeof(Microsoft.CSharp.RuntimeBinder.Binder).Assembly.Location));
        refs.Add(MetadataReference.CreateFromFile(typeof(System.Runtime.CompilerServices.DynamicAttribute).Assembly.Location));
        refs.Add(MetadataReference.CreateFromFile(typeof(System.Runtime.CompilerServices.CallSite).Assembly.Location));

        // "Batteries-included" reference set — common framework assemblies that
        // user .csx files reach for without thinking. Added because the script
        // host explicitly replaces ScriptOptions.Default.References via
        // WithReferences, so anything not enumerated here is genuinely
        // unresolvable inside a script. Each entry costs ~nothing per script
        // run; missing any of them costs a confused compile error and a
        // detour through `using/include?` debugging. See FR
        // `FR_scripting_dx_and_outlook_search.md` § Item 1.
        refs.Add(MetadataReference.CreateFromFile(typeof(System.Text.RegularExpressions.Regex).Assembly.Location));
        refs.Add(MetadataReference.CreateFromFile(typeof(System.Text.Json.JsonSerializer).Assembly.Location));
        refs.Add(MetadataReference.CreateFromFile(typeof(System.Net.Http.HttpClient).Assembly.Location));
        refs.Add(MetadataReference.CreateFromFile(typeof(System.Xml.XmlReader).Assembly.Location));
        refs.Add(MetadataReference.CreateFromFile(typeof(System.Xml.XmlDocument).Assembly.Location));
        refs.Add(MetadataReference.CreateFromFile(typeof(System.Diagnostics.Process).Assembly.Location));
        refs.Add(MetadataReference.CreateFromFile(typeof(System.Net.WebUtility).Assembly.Location));

        var options = ScriptOptions.Default
            .WithReferences(refs)
            .WithImports(new[]
            {
                "System",
                "System.Collections.Generic",
                "System.IO",
                "System.Linq",
                "System.Runtime.InteropServices",
            }.Concat(plugin.ScriptUsings))
            .WithFilePath(Path.GetFullPath(scriptPath))
            .WithEmitDebugInformation(true);

        // Redirect script Console.* into our output writer.
        var originalOut = Console.Out;
        var originalErr = Console.Error;
        Console.SetOut(output);
        Console.SetError(output);
        try
        {
            // Roslyn's internal scripting host creates its own AssemblyLoadContext
            // and loads referenced assemblies there. The plugin's own assembly is
            // already loaded in our PluginLoadContext, so without intervention
            // Roslyn loads a SECOND copy and the runtime globals object's Type
            // doesn't match the type Roslyn binds the script against — manifests
            // as "[A]ExcelGlobals cannot be cast to [B]ExcelGlobals". Telling
            // the InteractiveAssemblyLoader to reuse our already-loaded plugin
            // and globals assemblies fixes it.
            var loader = new InteractiveAssemblyLoader();
            loader.RegisterDependency(plugin.GetType().Assembly);
            loader.RegisterDependency(plugin.GlobalsType.Assembly);

            // Source-is-truth contract: the script file's bytes ARE what
            // Roslyn compiles. We never rewrite, prepend, or otherwise mutate
            // the source the author wrote — every reader (IDE, LLM, audit
            // tool, GitHub diff) sees exactly the code that ran. The
            // alternative (an injected preamble) was implemented in v0.4.2
            // and removed in v0.5.0; see CHANGELOG and the rejection note in
            // FR_office_script_interop_alias.md for the reasoning at
            // app-store scale.
            using var scriptStream = File.OpenRead(scriptPath);
            var script = CSharpScript.Create(scriptStream, options, plugin.GlobalsType, loader);
            var diags = script.Compile();
            var errors = diags.Where(d => d.Severity == DiagnosticSeverity.Error).ToList();
            if (errors.Count > 0)
            {
                foreach (var d in errors) output.WriteLine(AugmentOfficeDiagnostic(d.ToString()));
                return 3;
            }
            var state = await script.RunAsync(globals);
            if (state.Exception is not null)
            {
                output.WriteLine("SCRIPT EXCEPTION: " + state.Exception);
                return 4;
            }
            return 0;
        }
        catch (CompilationErrorException ex)
        {
            foreach (var d in ex.Diagnostics) output.WriteLine(d.ToString());
            return 3;
        }
        catch (Exception ex)
        {
            output.WriteLine("HOST EXCEPTION: " + ex);
            return 5;
        }
        finally
        {
            Console.SetOut(originalOut);
            Console.SetError(originalErr);
        }
    }

    /// <summary>
    /// Office-interop CS0104 collision pattern. Matches diagnostics like
    /// <c>error CS0104: 'Range' is an ambiguous reference between
    /// 'Microsoft.Office.Interop.Word.Range' and 'System.Range'</c> so we
    /// can append an actionable hint pointing at the conventional alias fix.
    /// </summary>
    private static readonly Regex OfficeCs0104Rx = new(
        @"error CS0104: '(?<name>\w+)' is an ambiguous reference between '(?<office>Microsoft\.Office\.Interop\.(?<app>Word|Excel|PowerPoint|Outlook))\.\w+' and 'System\.\w+'",
        RegexOptions.Compiled);

    /// <summary>
    /// Conventional two-letter alias per Office app. These are the same
    /// short names used throughout LLM/scripting.md and the per-plugin
    /// new-script templates, so the hint we emit lines up with the docs.
    /// </summary>
    private static readonly Dictionary<string, string> OfficeAliasSuggestions = new(StringComparer.Ordinal)
    {
        ["Word"]       = "Wd",
        ["Excel"]      = "Xl",
        ["PowerPoint"] = "Pp",
        ["Outlook"]    = "Ol",
    };

    /// <summary>
    /// If the diagnostic is a CS0104 Office-interop / BCL collision, append
    /// a short actionable hint to the error text pointing at the alias-
    /// declaration fix. The original diagnostic text — including its
    /// <c>(line,col)</c> span — is preserved verbatim; we only ADD a
    /// trailing hint line. Other diagnostics pass through unchanged.
    /// <para>
    /// This is the v0.5.0 replacement for the v0.4.2 ScriptUsingAliases
    /// preamble mechanism. Augmenting the error message keeps the script's
    /// source-is-truth contract intact (no host-side rewriting) while still
    /// turning a generic CS0104 into a one-step fix the author can apply
    /// without leaving their editor.
    /// </para>
    /// </summary>
    private static string AugmentOfficeDiagnostic(string diagText)
    {
        var m = OfficeCs0104Rx.Match(diagText);
        if (!m.Success) return diagText;
        string app = m.Groups["app"].Value;
        string name = m.Groups["name"].Value;
        string office = m.Groups["office"].Value;
        if (!OfficeAliasSuggestions.TryGetValue(app, out var alias))
            return diagText;
        return diagText
            + $"\n  -> Hint: add this to the top of your script:"
            + $"\n         using {alias} = global::{office};"
            + $"\n       then use '{alias}.{name}' instead of bare '{name}',"
            + $"\n       or qualify the BCL side as 'System.{name}'."
            + $"\n       See LLM/scripting.md for the full collision table.";
    }
}
