# combridge — LLM Reference Index

All paths in this folder are relative to the repo root (the directory
containing `combridge.sln`). No absolute paths — the docs are portable
across machines and checkouts.

TFM: `net10.0-windows`  |  PLATFORM: `x64`  |  LANG: `C# latest`
HOST_EXE: `combridge.exe`  |  ENTRY: `src/ComBridge.Cli/Program.cs::Main`
PLUGIN_LOAD_DIR: `<exeDir>/plugins/<Name>/<assembly>.dll`

## File map

**Bootstrap (read at session start):** this file → `LLM/workflow.md` → `LLM/api.md`.

| File | Purpose | Read when |
|---|---|---|
| `LLM/workflow.md` | **Task router.** "User wants X" → which docs to read in what order | After this file — always |
| `LLM/api.md` | Every public type + signature in `ComBridge.Core` + stability-tier table | Implementing a plugin or extending Core |
| `LLM/cli.md` | Exact CLI grammar, flag parsing order, exit codes | Writing scripts that call `combridge.exe` |
| `LLM/plugins.md` | Per-plugin specifics for the 5 shipped plugins (SW + Excel + Word + PowerPoint + Outlook) | Using or debugging an existing plugin |
| `LLM/scripting.md` | Per-plugin .csx recipe cookbook (Excel cell access, Word find/replace, Outlook mail, SW iteration, etc.) | Writing a Roslyn script |
| `LLM/extending.md` | Per-user / per-site `.csx` commands via `plugins/<Name>/commands/` auto-discovery. Plus "why not DLL-based sub-plugins?" criteria for the deferred Shape B. | Adding a custom named command to a plugin without forking the plugin |
| `LLM/authoring.md` | **Prescriptive guide for building a NEW plugin.** Four discovery patterns + worked examples for AutoCAD/Inventor/Acrobat/Visio/BricsCAD + verification checklist | Adding plugin support for an app NOT shipped in `src/plugins/` |
| `LLM/build.md` | Build pitfalls table (csproj, NuGet PIA, namespace, etc.) | A build error or runtime "type not found" |
| `LLM/paths.md` | Machine-specific path resolution chain (paths.props > env > registry > default > validation) | A plugin can't find its interop DLLs; adding a HintPath-based plugin |
| `LLM/troubleshooting.md` | Consolidated error → cause → fix catalog across all phases (build / load / attach / cast / script / COM runtime) | An error fires at any phase |
| `LLM/consuming.md` | Library-mode: a third-party CLI references `ComBridge.Core` directly. Stability tiers, attach patterns, multi-ProgID. | A tool needs ROT attach + session picking + scripting host but doesn't fit as a CLI invocation or as a plugin |
| `LLM/symbols.md` | Symbol → file index across the whole repo | Locating a class or method without grepping |

## Top-level facts

- One host exe, multiple plugins. Plugins live in `plugins/<Name>/` next to the exe.
- Plugins are loaded via per-folder `AssemblyLoadContext`. Default-context assemblies are preferred over plugin-local copies.
- ROT attach uses ole32 P/Invokes (`Marshal.GetActiveObject` removed from modern .NET). Types: `System.Runtime.InteropServices.ComTypes`.
- Multi-instance support: `list-sessions` built-in command + `--session <selector>` flag. Selector forms: `N` (1-based MRU index), `pid:NNNN`, `last`/`mru`/`recent` (explicit MRU keyword), or any other string (case-insensitive title substring).
- **Default session ordering: MRU (most-recently-used).** With no `--session`, the bridge picks the session whose window is highest in Windows' desktop Z-order — the one the user was last focused on. Implementation via `GetTopWindow` + `GetWindow(GW_HWNDNEXT)` Z-order walk in `SessionPicker.RankByZOrder`. Sessions whose process has no visible top-level window fall to the end (preserving ROT discovery order among themselves).
- Plugins referencing interop by file path resolve it via a 5-layer chain (`paths.props` > env var > Windows registry > default), validated at build time with `error COMBRIDGE001`. NuGet-PIA plugins (Excel) skip this entirely. Full spec: `LLM/paths.md`.
- Office plugins (Word/PowerPoint/Outlook in addition to Excel) reference their interop assemblies from the Office GAC at `C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.<App>\15.0.0.0__71e9bce111e9429c\` plus `office.dll` (Microsoft.Office.Core). All four Office apps share the same `_Application` dispinterface pattern — cast to `_Application`, NOT `Application`.
- COM attach uses two complementary paths: ROT-moniker walk via `RotMonikerPatterns` (multi-instance: SW `SolidWorks_PID_<pid>`, Office file-monikers `\.xlsx$`/`\.docx$`/`\.pptx$`) PLUS `oleaut32!GetActiveObject` fallback for class-moniker apps (Outlook). Plugin's `TryExtractRoot` hook ascends document RCWs to their parent Application. SessionPicker dedupes by PID and filters dead bindings (entries where the plugin's `DescribeInstance` returns BOTH null PID AND empty title — generic mechanism, applies to any COM host with transient sidecar processes, not just Office).
- Shared-instance shims (Office 365 being the prototypical case) consolidate per-process state back into a single host even when a new process was spawned. Multi-instance code paths are correct and verified against shim-free apps (SolidWorks). For shimmed apps, multi-instance is supported in code but rare in live use. Full notes: `C:\personal_rag\claude_code\lesson_20260521_office365_shared_instance_quirk.md`.
- **Three official consumption modes**: (1) call `combridge.exe` from a shell, (2) write a plugin DLL, (3) reference `ComBridge.Core.dll` as a library from a third-party tool. The library mode is for bespoke tools with their own UX or multi-ProgID workflows (SW + DocMgr). Stability tiers per type are listed at the top of `LLM/api.md`. Full guide: `LLM/consuming.md`.
- **Cross-platform as of v0.3.0**: `ComBridge.Core` multi-targets `net10.0` + `net10.0-windows`. The CLI binary similarly multi-targets. Plugins declare `IComBridgePlugin.SupportedPlatforms`; `PluginLoader` silently filters by the current OS. Windows plugins use COM via `RotHelper`/`SessionPicker` (#if WINDOWS-gated). Mac/Linux plugins use platform-native automation (e.g. `osascript` for AppleScript on macOS). The CLI contract (`combridge <app> <command>`) is identical across OSes — a ScripTree `.scriptree` file works without per-OS branching as long as a matching plugin exists for the target OS. First shipped Mac plugin: `ComBridge.Plugins.Excel.Mac` (v0.3.0).
- Roslyn scripting (`Microsoft.CodeAnalysis.CSharp.Scripting` 4.13.0) compiles user `.csx` against plugin globals + plugin-supplied references.
- Scripts must be valid Roslyn-script C# (top-level statements; `var` declarations are hoisted but assignments are not — referencing a top-level `var` before its declaration runs returns `null` at runtime, not a compile error).
- `using var x = ...` (C# 8 declaration form) is NOT supported in Roslyn scripting; use `using (var x = ...) { ... }` block form, or just `var x = ...` for one-shot scripts.

## Built-in commands (every plugin gets these for free)

| Command | Implementer | Effect |
|---|---|---|
| `list-plugins` | `ComBridge.Cli/Program.cs` | Prints discovered plugins (no plugin name required on CLI) |
| `<plugin> list-commands` | `ComBridge.Cli/Program.cs` | Prints commands for one plugin |
| `<plugin> list-sessions` | `ComBridge.Cli/Program.cs` (uses `SessionPicker.Enumerate`) | Lists running instances with index/PID/title |
| `<plugin> run-script <file.csx>` | `ComBridge.Core/Commands/RunScriptCommand.cs` | Compiles + runs script with plugin globals |

## Defaults reference

| Plugin | OS | ProgIDs / app name | AllowCreateNew | GlobalsType | Auto-imported namespaces |
|---|---|---|---|---|---|
| `solidworks` | Windows | `["SldWorks.Application"]` | `false` | `SwGlobals` | `SolidWorks.Interop.{sldworks, swconst, swcommands}` |
| `excel` (Windows plugin) | Windows | `["Excel.Application"]` | `true` | `ExcelGlobals` | `Microsoft.Office.Interop.Excel` |
| `word` | Windows | `["Word.Application"]` | `true` | `WdGlobals` | `Microsoft.Office.Interop.Word` |
| `powerpoint` | Windows | `["PowerPoint.Application"]` | `true` | `PptGlobals` | `Microsoft.Office.Interop.PowerPoint` |
| `outlook` | Windows | `["Outlook.Application"]` | `true` | `OlGlobals` | `Microsoft.Office.Interop.Outlook` |
| `excel` (Mac plugin) | macOS | `["Microsoft Excel"]` (AppleScript app name) | `true` | `XlMacGlobals` | `ComBridge.Plugins.Excel.Mac` |

Two plugins share `Name = "excel"` — the one targeting the current OS
loads, the other is silently filtered out by `PluginLoader` per its
`SupportedPlatforms`. So `combridge excel <command>` works the same on
Windows and macOS (different backend, same CLI contract).

## When verifying SolidWorks API calls

If a canonical SOLIDWORKS API reference is available on the host system
(typical filenames: `sldworks_methods_v3_llm.rag` for method signatures,
`swconst_enums.txt` for enum integer values), grep those files for the
exact `[IInterfaceName]` block before trusting any signature. Method
overloads and parameter ordering have changed across SW versions.
Without a canonical reference handy, assume signatures may be wrong and
verify via a small test before relying on them in batch code.
