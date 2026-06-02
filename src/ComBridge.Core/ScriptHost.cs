using System.Text;
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

            // Build the alias preamble (FR_office_script_interop_alias.md).
            // Each plugin-contributed alias is rendered as `using <alias>; `
            // and the whole set is concatenated onto a SINGLE first line so
            // line-number offsetting is exactly 1 (the script body's line 1
            // becomes compiled line 2). If a plugin contributes nothing, the
            // preamble is empty and there's zero behavior change.
            var aliases = plugin.ScriptUsingAliases?.ToList() ?? new List<string>();
            string preamble = aliases.Count == 0
                ? ""
                : string.Concat(aliases.Select(a => $"using {a}; ")) + "\n";
            int preambleLineOffset = preamble.Length == 0 ? 0 : 1;

            // To keep PDB emit happy (CS8055 needs explicit encoding), we
            // build a MemoryStream that preserves the source file's BOM (if
            // any) and uses the matching encoding for the preamble bytes.
            // Without this, mixing UTF-8-BOM script content with default-
            // encoded preamble would corrupt non-ASCII characters in the body.
            byte[] scriptBytes = File.ReadAllBytes(scriptPath);
            (Encoding enc, int bomLen) = DetectEncoding(scriptBytes);
            using var ms = new MemoryStream(bomLen + preamble.Length * 2 + scriptBytes.Length);
            if (bomLen > 0) ms.Write(scriptBytes, 0, bomLen);
            if (preamble.Length > 0)
            {
                byte[] preBytes = enc.GetBytes(preamble);
                ms.Write(preBytes, 0, preBytes.Length);
            }
            ms.Write(scriptBytes, bomLen, scriptBytes.Length - bomLen);
            ms.Position = 0;

            var script = CSharpScript.Create(ms, options, plugin.GlobalsType, loader);
            var diags = script.Compile();
            var errors = diags.Where(d => d.Severity == DiagnosticSeverity.Error).ToList();
            if (errors.Count > 0)
            {
                foreach (var d in errors) output.WriteLine(RemapDiagnosticLine(d.ToString(), preambleLineOffset));
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
    /// Detect the encoding + BOM length of raw source bytes. Default is
    /// UTF-8 without BOM (Roslyn's expectation for the streaming overload).
    /// </summary>
    private static (Encoding enc, int bomLen) DetectEncoding(byte[] bytes)
    {
        if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
            return (new UTF8Encoding(true), 3);
        if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
            return (Encoding.Unicode, 2);     // UTF-16 LE
        if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
            return (Encoding.BigEndianUnicode, 2);  // UTF-16 BE
        if (bytes.Length >= 4 && bytes[0] == 0 && bytes[1] == 0 && bytes[2] == 0xFE && bytes[3] == 0xFF)
            return (Encoding.UTF32, 4);
        return (new UTF8Encoding(false), 0);
    }

    /// <summary>
    /// Rewrite Roslyn diagnostic strings so reported line numbers point at
    /// the author's real source. Diagnostic text format is:
    /// <c>&lt;path&gt;(LINE,COL): severity ID: message</c> — we capture
    /// <c>(LINE,COL)</c> and subtract the preamble offset from LINE.
    /// <para>
    /// If the offending span is actually inside the injected preamble
    /// (which would be a bug in the plugin's alias declaration, not the
    /// author's script), we leave the line number alone so the bug is
    /// visible rather than presenting as "your script line 0" gibberish.
    /// </para>
    /// </summary>
    private static readonly Regex DiagLocRx = new(
        @"\((?<line>\d+),(?<col>\d+)\)",
        RegexOptions.Compiled);

    private static string RemapDiagnosticLine(string diagText, int preambleLineOffset)
    {
        if (preambleLineOffset == 0) return diagText;
        return DiagLocRx.Replace(diagText, m =>
        {
            int reportedLine = int.Parse(m.Groups["line"].Value);
            int col = int.Parse(m.Groups["col"].Value);
            int realLine = reportedLine - preambleLineOffset;
            if (realLine < 1) return m.Value;  // preamble-side error — leave as-is
            return $"({realLine},{col})";
        });
    }
}
