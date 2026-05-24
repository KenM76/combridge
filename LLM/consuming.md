# Consuming ComBridge.Core as a Library

`ComBridge.Core` is a public library. Third-party tools may reference it
directly to reuse ROT attach + session picking + Roslyn script host
without going through `combridge.exe` or writing a plugin.

This file is the LLM-optimized companion to `CONSUMING_CORE.md` (human).

## When library mode is the right consumption mode

| Symptom | Best mode |
|---|---|
| One-shot single-app command, output to file | CLI: `combridge <plugin> <cmd> out.txt` |
| Generic single-app automation, reusable from any .csx | New plugin (see `LLM/plugins.md`) |
| Bespoke tool with own UX (GUI/TUI), or orchestrates ≥2 COM servers in one process | **Library mode (this file)** |
| Stateless COM server (no concept of "session"; e.g. SwDocumentMgr) | Library mode + `RotHelper.AttachOrCreate` directly |

## Consumer csproj snippet (canonical)

```xml
<PropertyGroup>
  <ComBridgeRoot>$([System.Environment]::GetEnvironmentVariable('COMBRIDGE_ROOT'))</ComBridgeRoot>
  <ComBridgeRoot Condition="'$(ComBridgeRoot)' == ''">D:\Dev\combridge</ComBridgeRoot>
</PropertyGroup>
<ItemGroup>
  <Reference Include="ComBridge.Core">
    <HintPath>$(ComBridgeRoot)\src\ComBridge.Core\bin\Release\net10.0-windows\ComBridge.Core.dll</HintPath>
    <Private>true</Private>
  </Reference>
  <Reference Include="ComBridge.Plugins.SolidWorks">  <!-- optional -->
    <HintPath>$(ComBridgeRoot)\plugins\SolidWorks\ComBridge.Plugins.SolidWorks.dll</HintPath>
    <Private>true</Private>
  </Reference>
</ItemGroup>
<Import Project="$(ComBridgeRoot)\src\plugins\Common.Paths.props" />
```

`$(ComBridgeRoot)` resolution mirrors `LLM/paths.md`'s chain: `/p:` > env var > default. No new mechanism.

## Stability tiers (FR §3.1 acceptance criterion)

| Type / member | Tier | Breakability |
|---|---|---|
| `IComBridgePlugin` (interface) | **Stable** | New methods only via default impls (e.g. `DescribeInstance`). |
| `IBridgeCommand` (interface) | **Stable** | |
| `SessionInfo` (record) | **Stable** | New properties additive; rename/remove = break. |
| `RotHelper.TryGetActiveObject` | **Stable** | |
| `RotHelper.EnumerateActiveObjects` | **Stable** | |
| `RotHelper.AttachOrCreate` | **Stable** | |
| `SessionPicker.PidFromHwnd` | **Stable** | |
| `SessionPicker.Enumerate` | **Stable** | |
| `SessionPicker.Resolve` | **Stable** | Selector grammar in `LLM/cli.md` is part of the contract. |
| `ScriptHost.RunAsync(plugin, globals, scriptPath, output)` | **Stable** | Internal Roslyn invocation details are not. |
| `PluginLoader.LoadAll(pluginRoot?)` | **Stable** | |
| `PluginLoader.DefaultPluginRoot` | **Stable** | |
| Inner `PluginLoadContext` (private) | **Internal** | |
| `SolidWorksPlugin`, `ExcelPlugin` (concrete plugin types) | **Stable shape, evolving members** | Class names + ctor signatures stable; members may grow additively. |
| `SwGlobals`, `ExcelGlobals` (globals types) | **Stable shape, evolving members** | Field names + types stable; new fields may be added. |
| `ComBridge.Core.Commands.*` (e.g. `RunScriptCommand`) | **Internal-detail-but-public** | Used by CLI dispatch. External consumers: use `plugin.Commands` to find an `IBridgeCommand` by name instead. |

Rule of thumb: if it appears in `LLM/api.md`, it's stable. If it doesn't,
treat as internal even when the C# modifier says `public`.

## Attach patterns

### Single plugin (SW)

```csharp
var plugin   = new SolidWorksPlugin();
var sessions = SessionPicker.Enumerate(plugin);
var picked   = SessionPicker.Resolve(sessions, selectorOrNull)
               ?? throw new InvalidOperationException("No matching SW session.");
var globals  = (SwGlobals)plugin.CreateGlobals(picked.Root);
```

### Plugin loaded by name (avoids hard-referencing the plugin DLL)

```csharp
var comBridgeRoot = Environment.GetEnvironmentVariable("COMBRIDGE_ROOT") ?? @"D:\Dev\combridge";
var plugin = PluginLoader.LoadAll($@"{comBridgeRoot}\plugins")
                         .First(p => p.Name == "solidworks");
```

The two-layer C# resolution (`env var` → hardcoded default) mirrors the
MSBuild csproj snippet above. Same shape, same precedence.

### Stateless COM server (no plugin needed)

```csharp
var dmRoot = RotHelper.AttachOrCreate(
    new[] { "SwDocumentMgr.SwDocumentMgr" },
    createIfMissing: true);
// cast and use directly
```

### Multi-ProgID (SW + DocMgr in one process)

Combine the two above. Each COM root is independent; the consumer manages
lifetime. Don't try to express this as one `SessionPicker` call —
`SessionPicker` is plugin-scoped on purpose.

## `Common.Paths.props` from outside the tree

Import path: `$(ComBridgeRoot)\src\plugins\Common.Paths.props`.

What it provides to a consumer:
- The conditional `<Import>` of a local `paths.props` next to the consumer's `.csproj` (uses `$(MSBuildProjectDirectory)` so it picks the right one).
- The `ValidateRequiredInteropFiles` target keyed on `@(RequiredInteropFile)`. Same `COMBRIDGE001` error code.

Consumer declares its own `$(SolidWorksApiRedist)` (or equivalent) and
`@(RequiredInteropFile)` items exactly as a plugin does. See
`LLM/paths.md` § "Per-plugin csproj pattern" — the pattern is identical
for an external consumer.

Caveat: the path `src/plugins/Common.Paths.props` is current-tree-relative.
Once a NuGet package or `releases/` convention exists (FR §3.2), the
preferred import path will change. Until then, document the current
location in the consumer's README.

## Versioning — current state

| Mechanism | Status | Pin granularity |
|---|---|---|
| NuGet feed (local or hosted) | **planned**, not yet set up | semver |
| `releases/<version>/` folder | **planned**, not yet set up | folder name |
| `$(ComBridgeRoot)` + git commit hash | **only mechanism today** | commit |

Until NuGet/releases ships, consumers should document the combridge commit
hash they tested against in their own README. The library is stable
**within a commit**; build output paths, exe/assembly names, and folder
structure may shift between commits — anything a consumer's HintPath
touches is mutable until a versioned-release mechanism lands.

## Things deliberately NOT added (per FR §4)

| Rejected | Why |
|---|---|
| Multi-plugin `SessionPicker.Enumerate` | Stateless COM servers (DocMgr) have no sessions; adding ceremony for them is wrong. |
| C ABI / non-.NET surface | Out of scope. Different project if it comes up. |
| Cross-plugin command dispatch helper | `plugin.Commands.First(c => c.Name == "X").RunAsync(...)` is already trivial. |

## Files

| File | Audience | Role |
|---|---|---|
| `CONSUMING_CORE.md` (repo root) | humans | overview + canonical example |
| `LLM/consuming.md` | LLMs | this file — tables, exact snippets, stability tiers |
| `LLM/api.md` | LLMs | per-method signatures + the stability tier table (top of file) |
| `LLM/paths.md` § "External consumers" | LLMs | how an external `.csproj` imports `Common.Paths.props` |

## Quick decision tree

```
Is the work "use SW from a one-liner"?            → CLI (combridge solidworks ...)
Is the work "automate SW generically from any .csx"? → new plugin
Otherwise (custom UI, multi-app, long-lived, DocMgr involvement)? → library mode
```
