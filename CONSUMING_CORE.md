# Using `ComBridge.Core` as a library

`ComBridge.Core` is a **public library**, not just internal plumbing for
`combridge.exe`. If you're writing a tool that needs ROT attach, multi-instance
session picking, COM-app interop discovery, or the Roslyn script host —
reference `ComBridge.Core.dll` directly and skip the CLI.

This is the third officially-supported consumption mode:

| Mode | When it fits | When it doesn't |
|---|---|---|
| **Use the CLI** (`combridge solidworks <cmd>`) | One-shot, single-app commands. Output to a text/JSON file. | Tools with their own GUI, tools that need multiple COM servers in one process, long-lived services. |
| **Write a plugin** | The tool **is** "automate app X" itself — reusable from any user `.csx`. | The tool is a *consumer of* one or more plugins, not an *implementation of* one. |
| **Reference `ComBridge.Core` as a library** (this doc) | Custom CLI/service/GUI that orchestrates one or more COM apps and owns its own logic. Multi-ProgID workflows (SW + DocMgr, Excel + Outlook). | Trivial one-shot automation — use the CLI. Generic single-app automation — write a plugin. |

## Referencing it

In your tool's `.csproj`:

```xml
<PropertyGroup>
  <!-- Path to a built combridge tree. See "Resolving ComBridgeRoot" below. -->
  <ComBridgeRoot>$([System.Environment]::GetEnvironmentVariable('COMBRIDGE_ROOT'))</ComBridgeRoot>
  <ComBridgeRoot Condition="'$(ComBridgeRoot)' == ''">D:\Dev\combridge</ComBridgeRoot>
</PropertyGroup>

<ItemGroup>
  <Reference Include="ComBridge.Core">
    <HintPath>$(ComBridgeRoot)\src\ComBridge.Core\bin\Release\net10.0-windows\ComBridge.Core.dll</HintPath>
    <Private>true</Private>
  </Reference>
  <!-- Optional: reference the plugin DLL if you want to call SolidWorksPlugin directly
       (skip if you'd rather use PluginLoader.LoadAll to find it by name). -->
  <Reference Include="ComBridge.Plugins.SolidWorks">
    <HintPath>$(ComBridgeRoot)\plugins\SolidWorks\ComBridge.Plugins.SolidWorks.dll</HintPath>
    <Private>true</Private>
  </Reference>
</ItemGroup>

<!-- Reuse combridge's interop-path resolution chain (paths.props > env > registry > default)
     for any SolidWorks/AutoCAD/etc. interop DLLs your tool also needs directly. -->
<Import Project="$(ComBridgeRoot)\src\plugins\Common.Paths.props" />
```

## Minimal example — attach to SolidWorks and do something

```csharp
using ComBridge.Core;
using ComBridge.Plugins.SolidWorks;

var plugin   = new SolidWorksPlugin();
var sessions = SessionPicker.Enumerate(plugin);
var picked   = SessionPicker.Resolve(sessions, args.Length > 0 ? args[0] : null)
               ?? throw new InvalidOperationException("No matching SW session.");

var globals = (SwGlobals)plugin.CreateGlobals(picked.Root);
Console.WriteLine($"Attached to PID {picked.Info.Pid}: {picked.Info.Title}");
// globals.swApp, globals.swDoc, globals.swAssy, ... — your tool's code here.
```

If you don't want to hard-reference the plugin DLL, look it up by name instead:

```csharp
var comBridgeRoot = Environment.GetEnvironmentVariable("COMBRIDGE_ROOT") ?? @"D:\Dev\combridge";
var plugin = PluginLoader.LoadAll($@"{comBridgeRoot}\plugins")
                         .First(p => p.Name == "solidworks");
```

## Multi-ProgID example — SW *and* SwDocumentMgr in one process

`SessionPicker` is plugin-scoped (it asks "which running SW session?"). For
stateless COM servers like SwDocumentMgr that have no notion of a "session,"
just call `RotHelper.AttachOrCreate` directly:

```csharp
using ComBridge.Core;
using ComBridge.Plugins.SolidWorks;

// SolidWorks — attach to a specific running instance via the picker
var sw       = new SolidWorksPlugin();
var swRoot   = SessionPicker.Resolve(SessionPicker.Enumerate(sw), null)?.Root
               ?? throw new InvalidOperationException("No SW running.");
var swG      = (SwGlobals)sw.CreateGlobals(swRoot);

// SwDocumentMgr — stateless factory; just instantiate it
var dmRoot   = RotHelper.AttachOrCreate(new[] { "SwDocumentMgr.SwDocumentMgr" },
                                        createIfMissing: true);
// Cast and use dmRoot directly. DocMgr isn't a plugin concern; your tool owns it.
```

This is exactly the shape the `find-missing-refs` tool uses (live SW for
inspection + ReplaceReferencedDocument, DocMgr for reading GUIDs from closed
candidate files without opening them).

## Resolving `ComBridgeRoot`

The `.csproj` snippet above uses three layers (first non-empty wins). Pick
whatever fits your environment:

| Layer | How | When |
|---|---|---|
| 1. MSBuild override | `dotnet build /p:ComBridgeRoot=...` | CI / one-off |
| 2. Env var | `COMBRIDGE_ROOT` | shell / system / user-scope |
| 3. Default | hardcoded in csproj | dev convenience |

This mirrors the plugin path-resolution chain (`LLM/paths.md`) so anyone
who already knows how to configure combridge plugin paths knows how to
configure consumer references too.

## API stability

These are the types intended for external consumption. They have stable names
and contracts. The full per-method surface lives in `LLM/api.md`.

| Type | Stability | Notes |
|---|---|---|
| `interface IComBridgePlugin` | **Stable** | New methods only via default impls (like `DescribeInstance`) so existing consumers keep compiling. |
| `interface IBridgeCommand` | **Stable** | |
| `record SessionInfo` | **Stable** | Adding new properties is non-breaking; renaming or removing is. |
| `static class RotHelper` | **Stable** | `TryGetActiveObject`, `EnumerateActiveObjects`, `AttachOrCreate` are the contract. |
| `static class SessionPicker` | **Stable** | `Enumerate`, `Resolve`, `PidFromHwnd`. |
| `static class ScriptHost` | **Stable** | Only the `RunAsync(plugin, globals, scriptPath, output)` overload. Internal helpers may change. |
| `static class PluginLoader` | **Provisional** | `LoadAll` is stable. A `Describe` overload may be added; the inner `PluginLoadContext` is `private`. |
| Plugin types (`SolidWorksPlugin`, `SwGlobals`, `ExcelPlugin`, `ExcelGlobals`) | **Stable shape, evolving members** | Class names + globals field names won't change. Globals **may gain** new fields (additive). |
| Anything in `ComBridge.Core.Commands.*` | **Internal-detail-but-public** | Used by the CLI's dispatch. Don't depend on it externally; use `IBridgeCommand` from `plugin.Commands` instead. |

When in doubt: if it appears in `LLM/api.md`, treat it as stable. If it
doesn't, treat it as internal even when the C# modifier says `public`.

## Versioning (current state — known limitation)

Right now consumers pin to a `ComBridgeRoot` path. That works on one machine
but doesn't give you a real version. Two upgrade paths are on the roadmap:

- **Preferred:** publish `ComBridge.Core` (and the SolidWorks plugin) to a
  NuGet feed — local file-based feed is fine. Then consumers
  `<PackageReference Include="ComBridge.Core" Version="0.x.y" />`.
- **Interim:** a `releases/<version>/` folder in this repo with versioned
  DLL drops + a copy of `Common.Paths.props`.

Until one of those lands, the convention is: **pin to a known-good commit
of this repo**, and document the commit hash in your consumer's README.
The library *will* be stable within a commit, but build output paths,
exe/assembly names, and folder structure are not yet frozen and may change
between commits — anything a HintPath touches is a candidate for movement.
Once a `releases/<version>/` folder or NuGet feed exists, paths inside it
become the stable surface and HEAD layout is no longer the consumer's problem.

## Worked example — `find-missing-refs`

A real consumer should look like this (project layout):

```
D:\Dev\Projects\0003_SW_FindMissingFileByReference\
├── FindMissingRefs.csproj         ← references ComBridge.Core + Plugins.SolidWorks via HintPath
├── Program.cs                     ← argparse + main loop
├── ReferenceWalker.cs             ← uses SwGlobals.swAssy.GetComponents recursively
├── DocMgrLookup.cs                ← uses RotHelper.AttachOrCreate("SwDocumentMgr.SwDocumentMgr")
└── ResolutionUI.cs                ← the tool's own ConsoleSpectre / WPF / whatever
```

It does NOT have any of: a plugin manifest, a `combridge` subdirectory, a
`<plugins>` folder. It's a standalone CLI that happens to take a hard
dependency on `ComBridge.Core.dll`.

## When NOT to use library mode

- The whole tool is "I want `combridge solidworks active-doc` from a script" — just call the CLI.
- The tool is a *generic* automator for one app (e.g. "anything you'd want to do to SW") — write it as a plugin so other consumers and `.csx` users get the benefit.
- The tool fits in a one-off `.csx` script — use `combridge solidworks run-script foo.csx` instead.

Library mode is for **bespoke tools with their own UX and multi-app logic** that don't fit into either of the above.
