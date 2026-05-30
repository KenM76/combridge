# Changelog

All notable changes to this project will be documented in this file.
Format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/);
versions follow [SemVer](https://semver.org/spec/v2.0.0.html).

## [0.3.0] — cross-platform foundation (Windows + macOS)

### Added
- **Multi-targeted `ComBridge.Core`** — now builds for both `net10.0` and
  `net10.0-windows`. Windows-only code (`RotHelper`, `SessionPicker`
  Z-order/HWND helpers) is gated by `#if WINDOWS` per-method, so non-Windows
  plugins can reference Core without pulling in Win32 types.
- **`IComBridgePlugin.SupportedPlatforms`** — declares which OSes the
  plugin works on. `PluginLoader` silently skips plugins whose
  `SupportedPlatforms` doesn't include the current OS. Default = Windows
  only (matches v0.2.x plugin behavior; existing plugins keep working
  unchanged).
- **`IComBridgePlugin.FindSessions()`** — new default-interface method.
  Default impl on Windows delegates to `SessionPicker.Enumerate`
  (MRU-sorted via desktop Z-order). Non-Windows plugins MUST override
  with platform-native discovery (e.g. AppleScript on macOS).
- **`PluginLoader.IsSupportedOnCurrentOS(plugin)`** — public helper for
  checking platform support.
- **Multi-targeted `ComBridge.Cli`** — produces both a Windows binary
  (with full COM/ROT support) and a `net10.0` binary (for macOS/Linux
  with platform-neutral plugins only). Command dispatcher now routes
  all session discovery through `plugin.FindSessions()` so the CLI is
  OS-agnostic; `SessionPicker.Resolve` (cross-platform pure-string
  selector grammar) stays available on all OSes.
- **`ComBridge.Plugins.Excel.Mac` plugin** — first cross-platform plugin.
  Targets `net10.0`. Drives Microsoft Excel for Mac via `osascript`
  (AppleScript). Same CLI contract as the Windows Excel plugin
  (`combridge excel info`, `dump-sheet`, etc.) so a ScripTree `.scriptree`
  file targeting Excel works on both OSes without per-OS branching.

### Changed
- `SessionPicker` split into Windows-only methods (`PidFromHwnd`,
  `RankByZOrder`, `Enumerate`) and a cross-platform method (`Resolve`).
- `Program.cs` no-session-available fallback gated by `#if WINDOWS`;
  non-Windows builds emit a clear "no running session, open it manually"
  error rather than calling the Win32-only `RotHelper.AttachOrCreate`.

### Architecture
- Plugins are now categorized by platform:
  - **Windows-only**: `ComBridge.Plugins.{SolidWorks,Excel,Word,PowerPoint,Outlook}` (use COM, target `net10.0-windows`)
  - **macOS-only**: `ComBridge.Plugins.Excel.Mac` (uses `osascript`, targets `net10.0`)
  - Future: `Word.Mac`, `PowerPoint.Mac`, `LibreOffice` (any OS), etc.
- ScripTree files invoking `combridge <app> <command>` work uniformly
  on any OS where a plugin for that app exists — the CLI contract IS
  the cross-platform abstraction.

### Migration
- Existing Windows plugins keep working with zero source changes. They
  inherit `SupportedPlatforms => new[] { OSPlatform.Windows }` from the
  interface default.
- The `combridge.exe` binary for Windows is unchanged in behavior; all
  v0.2.0 commands, selectors, and scripts work identically.

## [0.2.0] — per-plugin `.csx` command extensions

### Added
- **Per-user / per-site scripted commands** — drop a `.csx` file in
  `plugins/<Name>/commands/` and `combridge` auto-discovers it as a
  named command (`combridge <plugin> <command-name>`). The script runs
  in the same Roslyn host as `run-script` with the plugin's globals
  available. See `LLM/extending.md`.
- `PluginLoader.GetScriptedCommands(plugin)` — public helper that
  enumerates the scripted commands for a given plugin.
- `Commands.ScriptedCommand` — public class wrapping one `.csx` file
  as an `IBridgeCommand`.
- `list-commands` output now labels commands by source: `(built-in)`,
  `(plugin)`, or `(script)`.

### Changed
- Command dispatcher in `Program.cs` now considers scripted commands
  after built-ins and typed plugin commands. Built-ins and typed
  plugin commands ALWAYS win on name collision — scripted commands
  can never shadow them.

### Deferred (not on roadmap)
- DLL-based sub-plugins ("Shape B"). Documented in `LLM/extending.md`
  with the specific scenarios that would warrant implementing it.

## [0.1.0] — initial public release

Generic COM-automation host for Windows desktop apps, with five shipped
plugins and a Roslyn `.csx` scripting host.

### Plugins
- **SolidWorks** (`solidworks`) — attach to running SLDWORKS.EXE via per-process `SolidWorks_PID_<pid>` ROT monikers. Multi-instance.
- **Excel** (`excel`) — attach via Workbook file-moniker + Application ascent, plus `oleaut32!GetActiveObject` fallback. Multi-instance per the code paths; Office 365 shared-instance shim limits live observation.
- **Word** (`word`) — file-moniker pattern + ascent, MRU-aware.
- **PowerPoint** (`powerpoint`) — file-moniker pattern + ascent.
- **Outlook** (`outlook`) — single MAPI session via `oleaut32!GetActiveObject`.

### Core features
- **Plugin architecture** — drop a DLL in `plugins/<Name>/` and it's discovered. Per-folder `AssemblyLoadContext` isolation; default-context assemblies (Core, BCL, Roslyn) reused across plugins.
- **Session picker** — `list-sessions` built-in + `--session N|pid:NNNN|<title>|last` selector. Default attach is MRU (most-recently-focused window via desktop Z-order). Sidecar/dead-binding filter drops transient Office shared-instance ghosts.
- **Roslyn script host** — `run-script <file.csx>`. `dynamic` supported. Script encoding handled via Stream overload (no CS8055 from BOM-less files). Plugin assemblies registered with `InteractiveAssemblyLoader` to avoid ALC identity mismatch.
- **Path resolution** — 5-layer chain (`paths.props` > env var > Windows registry > default) with `error COMBRIDGE001` build-time validation. Applies to plugins that reference interop via `<Reference HintPath>`.
- **Library mode** — `ComBridge.Core.dll` is a public library; third-party tools can reference it for ROT attach + session picking + scripting without going through `combridge.exe`. Stability tiers documented in `LLM/api.md`.

### Documentation
- **Human docs**: `README.md`, `PLUGIN_GUIDE.md`, `CONSUMING_CORE.md`.
- **LLM-optimized docs**: 11 files under `LLM/` covering API surface, CLI grammar, build pitfalls, path resolution, plugin authoring (with worked examples for AutoCAD/Inventor/Acrobat/Visio/BricsCAD), scripting recipes, troubleshooting catalog, library-mode usage, symbol index, and a task-router workflow file.
- **Examples**: 14 ready-to-run `.csx` scripts across all five plugins, with `examples/README.md` index.
- **In-source XML docs** on every public type in `ComBridge.Core`.

### Built on
- .NET 10 (TFM `net10.0-windows`)
- `Microsoft.CodeAnalysis.CSharp.Scripting` 4.13.0
- Office PIA assemblies (Excel via NuGet, Word/PowerPoint/Outlook via GAC HintPath)
- SOLIDWORKS interop assemblies (HintPath via `Common.Paths.props` chain)
