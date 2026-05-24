# ComBridge.Core — Public API Surface

All paths relative to `src/ComBridge.Core/` (under the repo root).

## Stability contract

Every type below is intended for external consumption (see `LLM/consuming.md`
for the library-mode story). Stability tiers:

| Tier | Meaning |
|---|---|
| **Stable** | Name + listed signatures will not change. New members may be added (additive). Default-interface-method additions are allowed (consumers continue to compile). |
| **Stable shape, evolving members** | Class/record name + ctor signatures stable. Fields/properties may grow additively; existing ones won't be renamed or removed. |
| **Provisional** | Likely to stabilize; may get additive helpers (e.g. `PluginLoader.Describe`). No breaking changes planned. |
| **Internal-detail-but-public** | C# modifier says `public` but contract is not stable. Don't depend on it from external consumers. |

| Type / member | Tier |
|---|---|
| `IComBridgePlugin` (interface, incl. `DescribeInstance` default impl) | Stable |
| `IBridgeCommand` (interface) | Stable |
| `SessionInfo` (record) | Stable |
| `RotHelper.TryGetActiveObject` / `EnumerateActiveObjects` / `AttachOrCreate` | Stable |
| `SessionPicker.PidFromHwnd` / `Enumerate` / `Resolve` (selector grammar from `LLM/cli.md` is part of the contract) | Stable |
| `ScriptHost.RunAsync(plugin, globals, scriptPath, output)` | Stable (internal Roslyn invocation details are not) |
| `PluginLoader.LoadAll` / `DefaultPluginRoot` | Stable |
| `PluginLoader.PluginLoadContext` (private nested type) | Internal |
| `SolidWorksPlugin`, `ExcelPlugin` (concrete plugin classes) | Stable shape, evolving members |
| `SwGlobals`, `ExcelGlobals` (script globals types) | Stable shape, evolving members |
| `ComBridge.Core.Commands.RunScriptCommand` and anything else under `Commands.*` | Internal-detail-but-public |

Rule of thumb: anything listed below in this file is part of the stable
surface. Anything that exists in source but isn't listed here should be
treated as internal even when the C# modifier disagrees.

## IComBridgePlugin (`IComBridgePlugin.cs`)

```csharp
public interface IComBridgePlugin
{
    string Name { get; }                                  // CLI dispatch name (lowercase, e.g. "solidworks")
    string Description { get; }                           // shown by `list-plugins`
    string[] ProgIds { get; }                             // ROT lookup + new-instance creation
    bool AllowCreateNew { get; }                          // true = launch fresh if ROT empty
    Type GlobalsType { get; }                             // Roslyn `globalsType` parameter
    object CreateGlobals(object comRoot);                 // factory for the script globals
    IEnumerable<MetadataReference> ScriptReferences { get; }  // interop DLLs for the compiler
    IEnumerable<string> ScriptUsings { get; }             // namespaces auto-imported in scripts
    IEnumerable<IBridgeCommand> Commands { get; }         // plugin-specific commands

    // Default impl returns (null, null). Override to enable PID + title in --session.
    (int? Pid, string? Title) DescribeInstance(object comRoot) => (null, null);
}
```

## IBridgeCommand (`IBridgeCommand.cs`)

```csharp
public interface IBridgeCommand
{
    string Name { get; }                                  // CLI sub-command name
    string Usage { get; }                                 // shown by `list-commands`
    Task<int> RunAsync(object comRoot, string[] args, TextWriter output);
}
```

`comRoot` is the already-attached/created COM root. Args do NOT include the plugin name, command name, or output file. Return value = process exit code.

## SessionInfo (`SessionInfo.cs`)

```csharp
public sealed record SessionInfo(int Index, int? Pid, string? Title, string Description);
```

`Index` is 1-based, assigned by `SessionPicker` in ROT enumeration order. `Description` is built from `Pid` + `Title` (see `SessionPicker.Enumerate`).

## RotHelper (`RotHelper.cs`, `[SupportedOSPlatform("windows")]`)

```csharp
public static class RotHelper
{
    public static object?              TryGetActiveObject(string progId);
    public static IEnumerable<object>  EnumerateActiveObjects(IEnumerable<string> progIds);
    public static object               AttachOrCreate(IEnumerable<string> progIds, bool createIfMissing);
}
```

- `EnumerateActiveObjects` walks ROT to completion, yielding every match. Caller dedupes (typically by PID — same process can register multiple monikers).
- `AttachOrCreate` returns first ROT match; if none and `createIfMissing`, calls `Activator.CreateInstance(Type.GetTypeFromProgID(...))` per ProgID.
- P/Invokes: `ole32!GetRunningObjectTable`, `ole32!CreateBindCtx`. Types `IRunningObjectTable`, `IBindCtx`, `IMoniker` from `System.Runtime.InteropServices.ComTypes`.

## SessionPicker (`SessionPicker.cs`)

```csharp
public static class SessionPicker
{
    public static int? PidFromHwnd(IntPtr hwnd);          // user32!GetWindowThreadProcessId
    public static List<(object Root, SessionInfo Info)> Enumerate(IComBridgePlugin plugin);
    public static (object Root, SessionInfo Info)? Resolve(
        List<(object Root, SessionInfo Info)> sessions, string? selector);
}
```

`Enumerate` calls `RotHelper.EnumerateActiveObjects(plugin.RotMonikerPatterns)`
+ `RotHelper.TryCoGetActiveObject(plugin.ProgIds)`, dedupes by PID, drops
dead bindings (both PID and title null), then **sorts by desktop Z-order
(most-recently-focused first)** via `RankByZOrder`. `SessionInfo.Index`
is assigned 1-based per the MRU-sorted order, so `Index == 1` is always
the most-recently-used session.

`Resolve` selector grammar:
| selector | match |
|---|---|
| `null` or `""` | most-recently-used session (`sessions[0]` after MRU sort) |
| `last` / `mru` / `recent` (case-insensitive) | same as null — explicit MRU keyword |
| pure digits | `Index == int.Parse(selector)` (1-based in MRU order) |
| `pid:NNNN` (case-insensitive prefix) | `Pid == int.Parse(rest)` |
| anything else | `Title.Contains(selector, OrdinalIgnoreCase)` OR `Description.Contains(...)` |

Returns first match, or `null` if no match.

## ScriptHost (`ScriptHost.cs`)

```csharp
public static class ScriptHost
{
    public static Task<int> RunAsync(
        IComBridgePlugin plugin,
        object globals,
        string scriptPath,
        TextWriter output);
}
```

Steps:
1. Read `scriptPath`. Missing → exit 2.
2. Build `ScriptOptions` with `plugin.ScriptReferences` + BCL trio (`object`, `Enumerable`, `Console` assemblies) + `plugin.GetType().Assembly` + `plugin.GlobalsType.Assembly`.
3. Auto-import: `System`, `System.Collections.Generic`, `System.IO`, `System.Linq`, `System.Runtime.InteropServices`, plus `plugin.ScriptUsings`.
4. Redirect `Console.Out` and `Console.Error` to the supplied writer for the duration.
5. `CSharpScript.Create(code, options, plugin.GlobalsType)`. Compile diagnostics with `Severity == Error` → exit 3.
6. `script.RunAsync(globals)`. `state.Exception != null` → exit 4. Other host exception → exit 5. Success → 0.

## PluginLoader (`PluginLoader.cs`)

```csharp
public static class PluginLoader
{
    public static string DefaultPluginRoot { get; }       // = AppContext.BaseDirectory + "/plugins"
    public static IReadOnlyList<IComBridgePlugin> LoadAll(string? pluginRoot = null);
}
```

For each subdirectory of `pluginRoot`:
1. Look for `ComBridge.Plugins.<dirName>.dll` (preferred) or any `*Plugin*.dll` fallback.
2. Create a `PluginLoadContext` (collectible: false) rooted at that folder.
3. `LoadFromAssemblyPath` then enumerate types implementing `IComBridgePlugin` and instantiate via parameterless ctor.

`PluginLoadContext.Load(AssemblyName)` resolution order:
1. Already-loaded assembly in `Default.Assemblies` (deduplicates Core, BCL, Roslyn).
2. `AssemblyDependencyResolver.ResolveAssemblyToPath`.
3. `<pluginDir>/<name>.dll`.
4. `null` (let Default context handle it).

## Commands.RunScriptCommand (`Commands/RunScriptCommand.cs`)

Built-in. Constructed by `Program.cs` per plugin when `commandName == "run-script"`. Args: `[<scriptFile.csx>]`. Calls `plugin.CreateGlobals(comRoot)` then `ScriptHost.RunAsync`.

## NuGet dependencies

- `Microsoft.CodeAnalysis.CSharp.Scripting` 4.13.0 (Core)
- `Microsoft.Office.Interop.Excel` 15.0.4795.1001 (Excel plugin only) — must be `<EmbedInteropTypes>false</EmbedInteropTypes>` plus `<CopyLocalLockFileAssemblies>true</CopyLocalLockFileAssemblies>`
