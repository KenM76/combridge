# Symbol → File Index

All paths relative to the repo root (the directory containing `combridge.sln`).

## Public types in ComBridge.Core

| Symbol | File |
|---|---|
| `interface IComBridgePlugin` | `src/ComBridge.Core/IComBridgePlugin.cs` |
| `interface IBridgeCommand` | `src/ComBridge.Core/IBridgeCommand.cs` |
| `record SessionInfo` | `src/ComBridge.Core/SessionInfo.cs` |
| `static class RotHelper` | `src/ComBridge.Core/RotHelper.cs` |
| `static class SessionPicker` | `src/ComBridge.Core/SessionPicker.cs` |
| `static class ScriptHost` | `src/ComBridge.Core/ScriptHost.cs` |
| `static class PluginLoader` | `src/ComBridge.Core/PluginLoader.cs` |
| `class Commands.RunScriptCommand` | `src/ComBridge.Core/Commands/RunScriptCommand.cs` |

## Internal types

| Symbol | File |
|---|---|
| `Program` (entry) | `src/ComBridge.Cli/Program.cs` |
| `PluginLoader.PluginLoadContext` (private) | `src/ComBridge.Core/PluginLoader.cs` |

## SolidWorks plugin types

| Symbol | File |
|---|---|
| `class SolidWorksPlugin` | `src/plugins/ComBridge.Plugins.SolidWorks/SolidWorksPlugin.cs` |
| `class SwGlobals` | `src/plugins/ComBridge.Plugins.SolidWorks/SolidWorksPlugin.cs` |
| `class ActiveDocCommand` | `src/plugins/ComBridge.Plugins.SolidWorks/SolidWorksPlugin.cs` |

## Excel plugin types

| Symbol | File |
|---|---|
| `class ExcelPlugin` | `src/plugins/ComBridge.Plugins.Excel/ExcelPlugin.cs` |
| `class ExcelGlobals` | `src/plugins/ComBridge.Plugins.Excel/ExcelPlugin.cs` |
| `class InfoCommand` | `src/plugins/ComBridge.Plugins.Excel/ExcelPlugin.cs` |
| `class DumpSheetCommand` | `src/plugins/ComBridge.Plugins.Excel/ExcelPlugin.cs` |

## Word plugin types

| Symbol | File |
|---|---|
| `class WordPlugin` | `src/plugins/ComBridge.Plugins.Word/WordPlugin.cs` |
| `class WdGlobals` | `src/plugins/ComBridge.Plugins.Word/WordPlugin.cs` |
| `class WdInfoCommand` | `src/plugins/ComBridge.Plugins.Word/WordPlugin.cs` |

## PowerPoint plugin types

| Symbol | File |
|---|---|
| `class PowerPointPlugin` | `src/plugins/ComBridge.Plugins.PowerPoint/PowerPointPlugin.cs` |
| `class PptGlobals` | `src/plugins/ComBridge.Plugins.PowerPoint/PowerPointPlugin.cs` |
| `class PptInfoCommand` | `src/plugins/ComBridge.Plugins.PowerPoint/PowerPointPlugin.cs` |

## Outlook plugin types

| Symbol | File |
|---|---|
| `class OutlookPlugin` | `src/plugins/ComBridge.Plugins.Outlook/OutlookPlugin.cs` |
| `class OlGlobals` | `src/plugins/ComBridge.Plugins.Outlook/OutlookPlugin.cs` |
| `class OlInfoCommand` | `src/plugins/ComBridge.Plugins.Outlook/OutlookPlugin.cs` |

## Build/config files

| File | Purpose |
|---|---|
| `combridge.sln` | Solution, 4 projects |
| `Directory.Build.props` | Shared TFM, nullability, implicit usings |
| `.gitignore` | bin/, obj/, *.user, .vs/, `paths.props` (live override; `.example` is tracked) |
| `src/plugins/Common.Paths.props` | Shared path-resolution import + `ValidateRequiredInteropFiles` target (emits `error COMBRIDGE001`) |
| `src/plugins/ComBridge.Plugins.SolidWorks/paths.props.example` | Committed template for the local (gitignored) `paths.props` |

## Plugin output deployment paths

| Source plugin | Deploys to |
|---|---|
| `src/plugins/ComBridge.Plugins.SolidWorks/` | `plugins/SolidWorks/` |
| `src/plugins/ComBridge.Plugins.Excel/` | `plugins/Excel/` |
| `src/plugins/ComBridge.Plugins.Word/` | `plugins/Word/` |
| `src/plugins/ComBridge.Plugins.PowerPoint/` | `plugins/PowerPoint/` |
| `src/plugins/ComBridge.Plugins.Outlook/` | `plugins/Outlook/` |

## Examples

| File | Demonstrates |
|---|---|
| `examples/sw_active_doc.csx` | Reading active doc title/path/type with SW globals |
| `examples/excel_dump_active_sheet.csx` | Iterating UsedRange with Excel globals |

(More recipes available in `LLM/scripting.md` — covers all 5 plugins.)

## Documentation

| File | Audience |
|---|---|
| `README.md` | Humans — overview + usage |
| `PLUGIN_GUIDE.md` | Humans — author guide for new plugins |
| `CONSUMING_CORE.md` | Humans — library-mode consumption (reference `ComBridge.Core.dll` from a third-party tool) |
| `LLM/README.md` | LLMs — entry point, file map |
| `LLM/api.md` | LLMs — public API surface of Core + stability tiers |
| `LLM/cli.md` | LLMs — CLI grammar + exit codes |
| `LLM/plugins.md` | LLMs — per-plugin specifics + new-plugin recipe |
| `LLM/build.md` | LLMs — pitfalls table + csproj boilerplate |
| `LLM/paths.md` | LLMs — machine-specific path resolution chain + COMBRIDGE001 |
| `LLM/consuming.md` | LLMs — library-mode tables, stability tiers, attach patterns, multi-ProgID |
| `LLM/scripting.md` | LLMs — per-plugin .csx recipe cookbook (all 5 shipped plugins) |
| `LLM/authoring.md` | LLMs — prescriptive guide for building new plugins (4 discovery patterns, worked examples for AutoCAD/Inventor/Acrobat/Visio/BricsCAD) |
| `LLM/troubleshooting.md` | LLMs — consolidated error → cause → fix catalog (build / load / attach / cast / script / COM runtime) |
| `LLM/workflow.md` | LLMs — task router; bootstrap-second |
| `LLM/symbols.md` | LLMs — this file |

## Quick navigation aliases

Primary entry: **`LLM/workflow.md`** has the full task router. Quick aliases:

- "Add a plugin for [app]" → `LLM/authoring.md` (the prescriptive guide); then `LLM/build.md` § "csproj boilerplate"; if HintPath interop, also `LLM/paths.md`
- "Write a .csx script for [app]" → `LLM/scripting.md` recipes + `LLM/plugins.md` § "<app> plugin"
- "Fix any error" → `LLM/troubleshooting.md`
- "Build error specifically" → `LLM/build.md` § "Pitfalls table"
- "Plugin can't find interop DLLs / error COMBRIDGE001" → `LLM/paths.md`
- "Change CLI behavior" → `LLM/cli.md` + `src/ComBridge.Cli/Program.cs`
- "Change session-picking logic" → `LLM/api.md` § "SessionPicker" + `src/ComBridge.Core/SessionPicker.cs`
- "Build a third-party tool that uses ComBridge.Core" → `LLM/consuming.md`
- "Is this API safe to depend on from outside?" → `LLM/api.md` § "Stability contract"
- "Verify a SOLIDWORKS API call" → If `C:\sw_api_docs\rag_optimized\sldworks_methods_v3_llm.rag` and `swconst_enums.txt` are present, grep those. Also check `C:\personal_rag\solidworks\` for empirical crash/quirk lessons.
