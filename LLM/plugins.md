# Plugin Implementations

## SolidWorks plugin

LOCATION: `src/plugins/ComBridge.Plugins.SolidWorks/`
ASSEMBLY: `ComBridge.Plugins.SolidWorks.dll`
DEPLOY: `plugins/SolidWorks/`
NAMESPACE: `ComBridge.Plugins.SolidWorks`
INTEROP SOURCE: file path (NOT NuGet) → uses the path-resolution chain.

### Machine-specific paths

| Property | Env var | Registry (Win) | Default |
|---|---|---|---|
| `SolidWorksApiRedist` | `SOLIDWORKS_API_REDIST` | `HKLM\SOFTWARE\SolidWorks\Setup\'Solidworks Folder'` + `\api\redist` | `C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\api\redist` |

`RequiredInteropFile`: `SolidWorks.Interop.{sldworks,swconst,swcommands}.dll` under `$(SolidWorksApiRedist)`. Missing → `error COMBRIDGE001`. Local override: copy `paths.props.example` → `paths.props`. Full chain: `LLM/paths.md`.

### IComBridgePlugin values

| Member | Value |
|---|---|
| `Name` | `"solidworks"` |
| `ProgIds` | `["SldWorks.Application"]` |
| `AllowCreateNew` | `false` (heavy app — never silent-launch) |
| `GlobalsType` | `SwGlobals` |
| `ScriptUsings` | `SolidWorks.Interop.{sldworks, swconst, swcommands}` |
| `ScriptReferences` | All `SolidWorks.Interop.*.dll` next to the plugin DLL |

### SwGlobals members (exposed to user .csx)

| Member | Type | Source | Notes |
|---|---|---|---|
| `swApp` | `ISldWorks` | `(ISldWorks)comRoot` | always non-null |
| `swDoc` | `IModelDoc2?` | `swApp.ActiveDoc as IModelDoc2` | null if no doc |
| `swPart` | `IPartDoc?` | `swDoc as IPartDoc` if `swDocType == swDocPART` | null otherwise |
| `swAssy` | `IAssemblyDoc?` | `swDoc as IAssemblyDoc` if `swDocType == swDocASSEMBLY` | null otherwise |
| `swDrawing` | `IDrawingDoc?` | `swDoc as IDrawingDoc` if `swDocType == swDocDRAWING` | null otherwise |
| `swDocType` | `swDocumentTypes_e` | `(swDocumentTypes_e)swDoc.GetType()` | `default` if no doc |

### Plugin commands

| Command | Args | Effect |
|---|---|---|
| `active-doc` | none | Prints active doc title, path, type. Smoke test. |

### DescribeInstance — SW API calls used

| Call | Signature | Notes |
|---|---|---|
| `ISldWorks.Frame()` | `object Frame()` | Cast result to `IFrame` |
| `IFrame.GetHWndx64()` | `long GetHWndx64()` | Use this on x64; `GetHWnd()` returns `int` and truncates |
| `IModelDoc2.GetTitle()` | `string GetTitle()` | |
| `IModelDoc2.GetPathName()` | `string GetPathName()` | Empty string for unsaved docs |
| `IModelDoc2.GetType()` | `int GetType()` | Cast to `swDocumentTypes_e` |
| `swDocumentTypes_e` | enum | `swDocPART=1`, `swDocASSEMBLY=2`, `swDocDRAWING=3` |

Verify these against the canonical SOLIDWORKS API reference for the SW
version installed on the target machine before extending — overload sets
do shift between releases.

### Implementation notes

- `as IPartDoc/IAssemblyDoc/IDrawingDoc` — guarded by a prior `GetType()` check on the doc, so the underlying object IS the right interface; cast normally succeeds. Generic caveat: the C# `as` operator can silently return null on some COM RCWs — if a cast is unreliable in your context, hard-cast `(IPartDoc)obj` and let it throw.
- `DescribeInstance` wraps everything in try/catch and returns `(null, null)` on any failure — instances mid-startup may not have a reachable Frame yet.

## Excel plugin

LOCATION: `src/plugins/ComBridge.Plugins.Excel/`
ASSEMBLY: `ComBridge.Plugins.Excel.dll`
DEPLOY: `plugins/Excel/`
NAMESPACE: `ComBridge.Plugins.Excel`
INTEROP SOURCE: NuGet PIA (`Microsoft.Office.Interop.Excel`) → NO path-resolution chain. No `paths.props`, no `Common.Paths.props` import, no `RequiredInteropFile`. NuGet resolves the DLL; csproj only needs `EmbedInteropTypes=false` + `CopyLocalLockFileAssemblies=true` (see `LLM/build.md`). Reference example for "plugin that does NOT use `LLM/paths.md`".
INTEROP_ALIAS: `using Xl = global::Microsoft.Office.Interop.Excel;` (NOT `Excel = ...` — that namespace shadows the alias)

### IComBridgePlugin values

| Member | Value |
|---|---|
| `Name` | `"excel"` |
| `ProgIds` | `["Excel.Application"]` |
| `AllowCreateNew` | `true` (Excel often skips ROT until a workbook is open) |
| `GlobalsType` | `ExcelGlobals` |
| `ScriptUsings` | `Microsoft.Office.Interop.Excel` |
| `ScriptReferences` | `Microsoft.Office.*.dll`, `office.dll`, `Microsoft.Vbe.*.dll` next to plugin |

### ExcelGlobals members

| Member | Type | Source | Notes |
|---|---|---|---|
| `xlApp` | `Xl.Application` | `(Xl.Application)comRoot` | always non-null |
| `xlBook` | `Xl.Workbook?` | `app.ActiveWorkbook` | null if none |
| `xlSheet` | `Xl.Worksheet?` | `app.ActiveSheet as Xl.Worksheet` | null if active sheet is a Chart |

### Plugin commands

| Command | Args | Effect |
|---|---|---|
| `info` | none | Prints version, visibility, active workbook/sheet names |
| `dump-sheet` | `[sheetName]` | TSV dump of `UsedRange.Value` for active or named sheet |

### DescribeInstance specifics

- HWND: `xlApp.Hwnd` (int property), cast to `IntPtr`
- Title: `xlApp.ActiveWorkbook?.Name`
- Wrap in try/catch; return `(null, null)` on any failure.

### Multi-instance caveat — shared-instance shims and transient sidecars

Code paths support multi-instance discovery (file-moniker walk +
`Document.Application` ascent + PID dedupe), and they're proven working
against apps without consolidation (SolidWorks). However: **some COM
hosts have shared-instance shims that consolidate state back into one
process**, leaving newly-spawned helper processes as transient sidecars.
Office 365 is the prototypical case — `New-Object -ComObject
Excel.Application` spawns a new EXCEL.EXE, but workbook state migrates
back to the existing Excel within seconds, leaving the new process with
no window (`MainWindowHandle=0`) and an unusable Application RCW.

`SessionPicker` generically handles this by dropping any entry where the
plugin's `DescribeInstance` returns both null PID AND empty title (the
dead-binding signature). Any future plugin whose target app shows similar
behavior automatically benefits.

To genuinely validate multi-instance code, test against an app that
doesn't consolidate (SolidWorks does not; AutoCAD does not; Outlook is
single-instance by design so it doesn't apply).

Full notes: `C:\personal_rag\claude_code\lesson_20260521_office365_shared_instance_quirk.md`.

## Word plugin

LOCATION: `src/plugins/ComBridge.Plugins.Word/`
ASSEMBLY: `ComBridge.Plugins.Word.dll`
DEPLOY: `plugins/Word/`
NAMESPACE: `ComBridge.Plugins.Word`
INTEROP SOURCE: GAC HintPath (`Microsoft.Office.Interop.Word.dll` + `office.dll`) — NOT NuGet.
INTEROP_ALIAS: `using Wd = global::Microsoft.Office.Interop.Word;`

| Member | Value |
|---|---|
| `Name` | `"word"` |
| `ProgIds` | `["Word.Application"]` |
| `AllowCreateNew` | `true` |
| `GlobalsType` | `WdGlobals` |
| `RotMonikerPatterns` | `[@"\.(docx|doc|docm|dotx|dotm|rtf)$"]` |
| `ScriptUsings` | `Microsoft.Office.Interop.Word` |

### WdGlobals members

| Member | Type | Source |
|---|---|---|
| `wdApp` | `Wd._Application` | `(Wd._Application)comRoot` |
| `wdDoc` | `Wd.Document?` | `app.ActiveDocument` |

### TryExtractRoot

Same pattern as Excel: bind file moniker → Document RCW → `.Application`
via dynamic dispatch.

### DescribeInstance

HWND via `app.ActiveWindow.Hwnd` with fallback to `app.Windows[1].Hwnd`
if no document open. Title from `app.ActiveDocument?.Name`.

### Commands

| Command | Args | Effect |
|---|---|---|
| `info` | none | Prints version, Visible, active document name, document count |

## PowerPoint plugin

LOCATION: `src/plugins/ComBridge.Plugins.PowerPoint/`
ASSEMBLY: `ComBridge.Plugins.PowerPoint.dll`
DEPLOY: `plugins/PowerPoint/`
NAMESPACE: `ComBridge.Plugins.PowerPoint`
INTEROP SOURCE: GAC HintPath.
INTEROP_ALIAS: `using Pp = global::Microsoft.Office.Interop.PowerPoint;`

| Member | Value |
|---|---|
| `Name` | `"powerpoint"` |
| `ProgIds` | `["PowerPoint.Application"]` |
| `AllowCreateNew` | `true` |
| `GlobalsType` | `PptGlobals` |
| `RotMonikerPatterns` | `[@"\.(pptx|ppt|pptm|potx|potm|ppsx|ppsm|pps|ppam)$"]` |
| `ScriptUsings` | `Microsoft.Office.Interop.PowerPoint` |

### PptGlobals members

| Member | Type | Source / Notes |
|---|---|---|
| `pptApp` | `Pp._Application` | `(Pp._Application)comRoot` |
| `pptPres` | `Pp.Presentation?` | `app.ActivePresentation` |
| `pptSlide` | `Pp.Slide?` | `app.ActiveWindow?.View?.Slide as Pp.Slide` — `View.Slide` returns `object` because SlideShowView has a different Slide type; cast required. |

### DescribeInstance

HWND via `app.HWND` (note uppercase, unlike Excel/Word). Title from
`app.ActivePresentation?.Name`.

### Commands

| Command | Args | Effect |
|---|---|---|
| `info` | none | Prints version, Visible, active presentation, presentation count, active slide index |

## Outlook plugin

LOCATION: `src/plugins/ComBridge.Plugins.Outlook/`
ASSEMBLY: `ComBridge.Plugins.Outlook.dll`
DEPLOY: `plugins/Outlook/`
NAMESPACE: `ComBridge.Plugins.Outlook`
INTEROP SOURCE: GAC HintPath.
INTEROP_ALIAS: `using Ol = global::Microsoft.Office.Interop.Outlook;`

| Member | Value |
|---|---|
| `Name` | `"outlook"` |
| `ProgIds` | `["Outlook.Application"]` |
| `AllowCreateNew` | `true` |
| `GlobalsType` | `OlGlobals` |
| `RotMonikerPatterns` | `[]` (empty — Outlook is single-instance MAPI; rely on `GetActiveObject` only) |
| `ScriptUsings` | `Microsoft.Office.Interop.Outlook` |

### OlGlobals members

| Member | Type | Source |
|---|---|---|
| `olApp` | `Ol._Application` | `(Ol._Application)comRoot` |
| `olNs` | `Ol.NameSpace` | `app.GetNamespace("MAPI")` — the only namespace Outlook supports |
| `olExplorer` | `Ol.Explorer?` | `app.ActiveExplorer()` |

### Why no `RotMonikerPatterns`

Outlook is fundamentally different from document-based Office apps. It's
one MAPI session per user, single process (`OUTLOOK.EXE`), no per-document
ROT registration. Attach uses `oleaut32!GetActiveObject("Outlook.Application")`
exclusively, which works reliably for Outlook because it DOES register a
class moniker — unlike Excel.

### DescribeInstance

PID via `Process.GetProcessesByName("OUTLOOK")[0].Id` (single-instance, so
the first match is correct). Title from `app.ActiveExplorer().CurrentFolder.Name`
(typically `"Inbox"`) with fallback to `Explorer.Caption`.

### Commands

| Command | Args | Effect |
|---|---|---|
| `info` | none | Prints version, current user name, Inbox item count, all stores (mail accounts), active explorer caption + current folder |

### Outlook scripting hints (Windows MAPI)

The `NameSpace` is where everything lives:

```csharp
// list all calendars
var calendars = olNs.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
foreach (Folder sub in calendars.Folders) { ... }

// iterate Inbox
var inbox = olNs.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
foreach (object item in inbox.Items)
{
    if (item is MailItem mail) { Console.WriteLine(mail.Subject); }
}
```

The `Stores` collection on `NameSpace` enumerates accounts (one entry per
configured mailbox / PST / OST).

## Recipe for adding a new plugin (mechanical)

```
1. Create dir: src/plugins/ComBridge.Plugins.<Name>/
2. Create <Name>Plugin.csproj  ← copy from PLUGIN_GUIDE.md template
3. Create <Name>Plugin.cs       ← implement IComBridgePlugin
4. Add project to combridge.sln (4 lines: Project entry + 4 ProjectConfigurationPlatforms entries)
5. dotnet build -c Release
6. Verify plugins/<Name>/<assembly>.dll exists + interop DLLs present
7. combridge list-plugins             ← should show the new plugin
8. combridge <Name> list-commands     ← should show built-ins + any custom
```

### Required IComBridgePlugin members

```csharp
string Name { get; }                                // lowercase
string Description { get; }
string[] ProgIds { get; }                           // priority order
bool AllowCreateNew { get; }
Type GlobalsType { get; }
object CreateGlobals(object comRoot);
IEnumerable<MetadataReference> ScriptReferences { get; }
IEnumerable<string> ScriptUsings { get; }
IEnumerable<IBridgeCommand> Commands { get; }       // can be empty array
```

### Optional override

```csharp
(int? Pid, string? Title) DescribeInstance(object comRoot);
```

**Strongly recommended.** If you skip it, sessions still appear in
`list-sessions` (entries with both null PID and empty title are dropped
by the dead-binding filter — so a plugin that returns `(null, null)`
universally yields an empty list), but:

- `--session pid:NNNN` and `--session "<title>"` selectors won't match
  (need PID and title respectively)
- **Z-order MRU sorting needs the PID** to look up each session's
  position in the desktop Z-order — without it, sessions sort to the
  end and the "most recent" guarantee degrades to ROT discovery order
- `list-sessions` rows show `(no info)` instead of `pid=… title=…`

Returning at least the PID is what unlocks the MRU default. Returning
title too is needed for the substring selector. Both are cheap to
extract for most apps — see the HWND properties table below.

### App-specific HWND properties (known)

| App | HWND access | Notes |
|---|---|---|
| SolidWorks | `((IFrame)swApp.Frame()).GetHWndx64()` returns `long` | Use `GetHWndx64`, not `GetHWnd` |
| Excel | `xlApp.Hwnd` returns `int` | |
| Word | `wdApp.ActiveWindow.Hwnd` returns `int` | (untested in this repo) |
| AutoCAD | `acadApp.HWND` returns int (handle) | (untested in this repo) |
| Inventor | `invApp.MainFrameHWND` returns `int` | (untested in this repo) |

In all cases pass to `SessionPicker.PidFromHwnd((IntPtr)hwnd)` which calls `user32!GetWindowThreadProcessId`.

# macOS plugins (Excel.Mac / Word.Mac / PowerPoint.Mac / Outlook.Mac)

Four AppleScript-backed Office plugins ship alongside the Windows ones,
sharing CLI names so `combridge <app> <command>` works the same on
Windows and macOS. `PluginLoader` filters by current OS via
`SupportedPlatforms`, so only one "excel" / "word" / "powerpoint" /
"outlook" plugin loads per machine.

## Shape they all share

- `TargetFramework`: `net10.0` (NOT `net10.0-windows`)
- `SupportedPlatforms`: `[OSPlatform.OSX]`
- Reference `ComBridge.Core` AND `ComBridge.Mac.Common` (shared
  `Osascript` helper — wraps `Process.Start("osascript", ...)`)
- Override `FindSessions()` with native discovery via
  `tell application "System Events" to (name of processes) contains "..."`
- `DescribeInstance` is stubbed (returns null/null) — `FindSessions`
  populates the SessionInfo directly without going through SessionPicker
- `CreateGlobals(comRoot)` ignores the `comRoot` (sentinel `new object()`
  from `FindSessions`) and constructs a fresh wrapper

## Per-plugin specifics

| Plugin | AppleScript app name | Globals | Built-in commands |
|---|---|---|---|
| `excel` (Mac) | `"Microsoft Excel"` | `xlApp` (`XlMacApp`) | `info`, `dump-sheet` |
| `word` (Mac) | `"Microsoft Word"` | `wdApp` (`WdMacApp`) | `info`, `extract-text`, `doc-stats` |
| `powerpoint` (Mac) | `"Microsoft PowerPoint"` | `pptApp` (`PptMacApp`) | `info`, `list-slides` |
| `outlook` (Mac) | `"Microsoft Outlook"` | `olApp` (`OlMacApp`) | `info`, `list-accounts` |

## What's different vs the Windows plugins (script-author level)

- **Globals types differ.** Windows `xlApp` is `Microsoft.Office.Interop.Excel._Application`
  with ~3000 members. Mac `xlApp` is `XlMacApp` with ~10 high-level
  helpers (Version, ActiveWorkbookName, GetCellValue, DumpUsedRange,
  etc.). A .csx using Windows-specific idioms (`xlApp.Workbooks[1].Worksheets[2].Cells[3,4].Value`)
  won't compile against the Mac plugin's globals.
- **Per-cell access is slow** — each call shells out to osascript
  (~5-20ms cold). Bulk operations (`DumpUsedRange`) batch into one
  invocation; use those instead of looping individual cells.
- **No "multi-instance" Office.** Mac Office is single-instance per
  user. `list-sessions` always returns 0 or 1 entries.
- **Outlook for Mac has a thinner dictionary** than Windows MAPI. No
  Stores collection, no Restrict() with Jet syntax, limited message
  body access. The plugin exposes accounts + inbox counts; richer
  scripting requires falling back to direct osascript.

## Implementation notes for future Mac plugin authors

- **`Microsoft Excel` vs `Excel`**: AppleScript app names use the full
  Office-branded name (`"Microsoft Excel"`, `"Microsoft Word"`, etc.),
  NOT just the bare app name. Get it wrong and AppleScript throws
  "Application isn't running" even when it is.
- **App-running probe**: `tell application "System Events" to (name of
  processes) contains "..."` returns `true`/`false`. Use `Osascript.TryRun`
  to handle failures silently.
- **PID extraction**: `tell application "System Events" to unix id of
  (first process whose name is "...")` returns the PID as a string;
  parse to int.
- **Active document title**: each Office app has its own AppleScript
  vocabulary (`active workbook`, `active document`, `active presentation`).
  See per-plugin source for the right one.
- **AppleScript string-list coercion**: `name of every X as text`
  returns a comma-space-separated list. Split on `", "` to recover
  individual items (with a small risk of false splits if a name itself
  contains `", "`).
- **Batching multi-call AppleScript**: when iterating something, wrap
  the loop inside one big AppleScript that emits a delimiter-joined
  result. A single osascript invocation is ~5-20ms cold; 100 cell-by-cell
  invocations is 0.5-2s.

## Known limitation — "New Outlook for Mac"

Microsoft's Catalyst-based "New Outlook for Mac" (rolling out 2024+)
further restricts AppleScript automation. This plugin targets the
classic Outlook for Mac. Migration to a thin REST wrapper over
Microsoft Graph may be necessary in a future release if Microsoft
removes AppleScript support entirely.
