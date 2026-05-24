# Troubleshooting

Consolidated error → cause → fix catalog for combridge. Errors are
grouped by phase (build / load / attach / runtime) so you can jump to
the right section based on when the error fires.

## Build-time errors

### `error COMBRIDGE001: ... could not locate its required interop assemblies`

**Cause.** A plugin declared `@(RequiredInteropFile)` items in its csproj
and at least one of those files doesn't exist on disk. The path was
resolved through `Common.Paths.props`'s chain (paths.props → env var →
registry → default) and none yielded a valid file.

**Fix.** Read the error's "Resolution chain (first non-empty wins)" block.
Set one of the layers correctly:
- `paths.props` (copy `paths.props.example` next to the .csproj and edit)
- env var (`SOLIDWORKS_API_REDIST` for SW, etc.)
- `/p:Property=value` on the build command line

See `LLM/paths.md` for the full chain.

### `error CS0234: type or namespace 'Application' does not exist in namespace 'ComBridge.Plugins.Excel'`

**Cause.** Namespace shadowing. The plugin's namespace ends in `.Excel`
(or `.Word`/`.Outlook`), and `using Excel = Microsoft.Office.Interop.Excel`
resolves `Excel.Application` against the LOCAL namespace, not the alias.

**Fix.** Alias to a non-clashing name:

```csharp
using Xl = global::Microsoft.Office.Interop.Excel;
using Wd = global::Microsoft.Office.Interop.Word;
using Pp = global::Microsoft.Office.Interop.PowerPoint;
using Ol = global::Microsoft.Office.Interop.Outlook;
```

Two-letter aliases (`Xl`/`Wd`/`Pp`/`Ol`) are conventional in combridge.

### `error CS0246: type or namespace 'IRunningObjectTable' not found`

**Cause.** `IRunningObjectTable`, `IBindCtx`, `IMoniker` live in
`System.Runtime.InteropServices.ComTypes`, NOT `System.Runtime.InteropServices`.

**Fix.** `using System.Runtime.InteropServices.ComTypes;`

### NuGet PIA package loads, but interop types resolve to embedded types

**Cause.** NuGet PIA packages default to `EmbedInteropTypes=true`. The
compiler inlines a stripped-down copy of the types into your assembly,
and they don't share identity with the same types loaded from the
real PIA at runtime. Symptoms: weird casting errors at runtime.

**Fix.**

```xml
<PackageReference Include="Microsoft.Office.Interop.Excel" Version="...">
  <EmbedInteropTypes>false</EmbedInteropTypes>
</PackageReference>
```

### Interop DLL doesn't appear in plugin output folder

**Cause.** NuGet transitive deps aren't copied to the output by default
for library projects.

**Fix.** Add to the plugin csproj `<PropertyGroup>`:

```xml
<CopyLocalLockFileAssemblies>true</CopyLocalLockFileAssemblies>
```

Excel plugin csproj is the reference example.

## Plugin-load errors

### `error: System.IO.FileNotFoundException: Could not load file or assembly 'office, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c'`

**Cause.** `office.dll` (Microsoft.Office.Core PIA) is required by every
Office interop assembly (Excel/Word/PowerPoint/Outlook), but NuGet's
`Microsoft.Office.Interop.*` packages don't pull it as a transitive
dependency. The first cast to `_Application` triggers type load and
the runtime can't find `office.dll`.

**Fix.** Add an explicit reference to the GAC copy:

```xml
<Reference Include="office">
  <HintPath>C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\office.dll</HintPath>
  <EmbedInteropTypes>false</EmbedInteropTypes>
  <Private>true</Private>
</Reference>
```

Office's installer puts `office.dll` in the GAC on any Office-installed
machine; the hardcoded path is portable.

Full notes: `C:\personal_rag\claude_code\lesson_20260521_excel_dispinterface_vs_coclass_cast.md`.

### `(no plugins discovered)` from `list-plugins`

**Cause.** One of:
- `combridge.exe` not staged next to a `plugins/` directory
- Plugin DLL doesn't follow naming convention `ComBridge.Plugins.<Name>.dll`
- Plugin DLL is in the right folder but `<Name>` doesn't match the parent directory name
- Plugin throws during instantiation (parameterless ctor failure)

**Fix.** Verify:
```powershell
ls .\plugins\<Name>\ComBridge.Plugins.<Name>.dll   # must exist with this exact name
.\combridge.exe list-plugins
```

If the file is there but doesn't load, check for ctor exceptions by
running with extra logging. (No built-in verbose flag exists yet; if
needed, instrument `PluginLoader.LoadAll` temporarily.)

## Attach / discovery errors

### `(no running <plugin> sessions in the ROT)` when the app IS running

**Cause.** `RotMonikerPatterns` doesn't match what's actually in the ROT.
Substring/regex mismatch.

**Diagnostic** — dump the live ROT to see what's there:

```csharp
// in a .csx, against ANY working plugin (e.g. excel):
using System.Runtime.InteropServices.ComTypes;
[DllImport("ole32.dll")] static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable rot);
[DllImport("ole32.dll")] static extern int CreateBindCtx(int reserved, out IBindCtx ctx);
GetRunningObjectTable(0, out var rot);
CreateBindCtx(0, out var bc);
rot.EnumRunning(out var en); en.Reset();
var mon = new IMoniker[1];
while (en.Next(1, mon, IntPtr.Zero) == 0) {
    mon[0].GetDisplayName(bc, null, out string dn);
    Console.WriteLine(dn);
}
```

Look at the actual display names. Common patterns:
- Class moniker: `!{CLSID-in-braces}` or `!ProgID`
- File moniker: full path to a document
- Custom moniker: vendor-specific format (`SolidWorks_PID_<pid>`)

Update `RotMonikerPatterns` to match the real format. Always anchor with
`^…$` to avoid false positives.

Full notes: `C:\personal_rag\claude_code\lesson_20260521_rot_walk_vs_getactiveobject.md`.

### `(no running <plugin> sessions in the ROT)` for Excel/Word when only an unsaved doc is open

**Cause.** Unsaved (no file path) documents don't register as file
monikers in the ROT. `oleaut32!GetActiveObject` is the fallback for this
case — it should still find the app even with no saved docs.

**Fix.** Verify your plugin's `ProgIds` is correct (the canonical
ProgID string), since `GetActiveObject` keys off it. For Excel:
`"Excel.Application"` exactly (not `"Xl.Application"` or anything else).

### Default attach picks the "wrong" session (multi-instance)

**Symptom.** With multiple instances of the app running (e.g. two
SOLIDWORKS sessions), `combridge <plugin> <command>` attaches to a
session the user wasn't actually working in.

**Cause.** As of the MRU change, the default with no `--session` is
*most-recently-focused* — the window highest in Windows' desktop
Z-order, excluding the terminal/IDE the user typed `combridge` into.
If the user switched apps in between (e.g. clicked SW window B, then
opened Notepad, then ran `combridge`), Notepad is at the top of
Z-order and SW window B is below it — but B still wins among matching
SW PIDs. So usually MRU does what you want.

**When MRU disagrees with intuition:**

1. **User restored a minimized window via taskbar** — that updates
   Z-order, so it counts as "most recent."
2. **User has not focused any matching window since boot** — Z-order
   defaults to launch order; MRU degrades to "approximately first
   launched."
3. **Window is minimized to system tray (no top-level window)** —
   falls to the end of the MRU sort; might lose to an instance the
   user hasn't actually touched recently.

**Fix.** Use a deterministic selector instead of relying on default:
- `--session pid:NNNN` — most reliable; PID never changes for a
  running process.
- `--session "<title-substring>"` — works if document titles are
  distinct.
- `--session N` — index in MRU order; only deterministic if you've
  just looked at `list-sessions` output.

If you want the LEGACY behavior (always attach to ROT-first instead
of MRU), there is currently no built-in selector for it. Use
`--session pid:NNNN` with an explicit PID, or write a small
library-mode tool that calls `RotHelper.EnumerateActiveObjects`
directly without MRU sorting.

### `--session pid:NNNN` says "matched no running instance" but the PID is real

**Cause.** Plugin's `DescribeInstance` isn't returning the PID
(returning `(null, title)` or `(null, null)` for the matched entry).

**Fix.** Verify the HWND extraction path. For each app, the property
name varies:

| App | HWND property |
|---|---|
| Excel | `xlApp.Hwnd` (lowercase) |
| Word | `wdApp.ActiveWindow?.Hwnd` (may be null if no doc) |
| PowerPoint | `pptApp.HWND` (UPPERCASE) |
| Outlook | not directly exposed — use `Process.GetProcessesByName("OUTLOOK")[0].Id` |
| SolidWorks | `((IFrame)swApp.Frame()).GetHWndx64()` — long, not int |

Wrap in try/catch; return `(pid, title)` even if one is null.

### List-sessions shows `(no info)` entries

**Cause.** `DescribeInstance` returned both null PID AND empty title.
SessionPicker drops these by default, but if you're seeing them, the
plugin returned at least one signal that wasn't usable.

**Fix.** Almost certainly Office's shared-instance shim handing off a
moniker between processes. Read the next section.

## Office shared-instance behavior

### `New-Object -ComObject Excel.Application` spawns a process but my new workbook ends up in the old Excel

**Cause.** Office 365 has a shared-instance shim. The new EXCEL.EXE
process briefly hosts the workbook, then Office routes the file open to
the existing user-interactive Excel for state consolidation. The new
EXCEL.EXE is left as a hollow sidecar with no window
(`MainWindowHandle=0`).

**Fix.** This is by design in modern Office. Workarounds (none reliable
on stock Office 365):
- Use the `/x` flag on `EXCEL.EXE` (mostly deprecated)
- Registry override `HKCU\SOFTWARE\Microsoft\Office\<ver>\Excel\Options\DisableMergeInstance` (build-dependent)
- Older Office (≤2013) installs honor `/x`

For combridge testing: validate multi-instance code paths against apps
without a shared-instance shim (SolidWorks, AutoCAD, BricsCAD).

Full notes: `C:\personal_rag\claude_code\lesson_20260521_office365_shared_instance_quirk.md`.

## Cast errors

### `System.InvalidCastException: Unable to cast COM object of type 'System.__ComObject' to interface type 'Microsoft.Office.Interop.Excel._Application' ... QueryInterface for IID '{000208D5-…}' failed ... E_NOINTERFACE`

**Cause.** Casting an Excel-like RCW to the wrong type. The COM object
returned from ROT or GetActiveObject implements the `_Application`
dispinterface (IID `{000208D5-…}`), NOT the `Application` co-class
(different RCW identity).

**Fix.** Cast to the dispinterface (note the underscore):

```csharp
var app = (Xl._Application)comRoot;   // ✓
var app = (Xl.Application)comRoot;     // ✗ E_NOINTERFACE
```

The same applies to all Office apps and many other COM dispatch types.
Look for the `_` prefix in the interop's interface list.

### `[A]Globals cannot be cast to [B]Globals` from script runtime

**Cause.** Roslyn's internal scripting host loaded the plugin assembly
in its own AssemblyLoadContext, while the actual globals object was
created by combridge's PluginLoadContext. Same DLL, two ALCs, two
distinct Types.

**Fix.** Already handled in `ScriptHost.cs`:

```csharp
var loader = new InteractiveAssemblyLoader();
loader.RegisterDependency(plugin.GetType().Assembly);
loader.RegisterDependency(plugin.GlobalsType.Assembly);
var script = CSharpScript.Create(stream, options, plugin.GlobalsType, loader);
```

If you see this error, ScriptHost has regressed — investigate. Full
notes: `C:\personal_rag\claude_code\lesson_20260521_roslyn_scripting_in_plugin_alc.md`.

## Script-time errors

### `error CS8055: Cannot emit debug information for a source text without encoding`

**Cause.** Script file was loaded as a `string` (no encoding metadata)
and `ScriptOptions.WithEmitDebugInformation(true)` is set. Roslyn needs
to know the source encoding to emit PDB info.

**Fix.** Already handled — `ScriptHost.cs` uses the Stream overload of
`CSharpScript.Create` which detects BOM or defaults to UTF-8.

If you see this error, the host has regressed to a string-based load.

### `error CS0656: Missing compiler required member 'Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo.Create'`

**Cause.** Script uses `dynamic` but `Microsoft.CSharp.dll` isn't in the
ScriptOptions references.

**Fix.** Already handled — `ScriptHost.cs` adds `Microsoft.CSharp` +
`DynamicAttribute` + `CallSite` refs to the default set. If you see
this error, the host has regressed.

### `error CS1980: ... 'System.Runtime.CompilerServices.DynamicAttribute' cannot be found`

**Cause.** Same family as CS0656 — the runtime types for `dynamic`
aren't all referenced. CS1980 specifically means `DynamicAttribute` is
missing.

**Fix.** Same as CS0656.

### `error CS1002: ; expected` at `using var x = ...`

**Cause.** Roslyn scripting host doesn't support the C# 8 `using var`
DECLARATION form, even though the language version compiles it
normally.

**Fix.** Use the classic `using` STATEMENT form:

```csharp
// ✗ doesn't work in .csx
using var stream = File.OpenRead(path);

// ✓ works
using (var stream = File.OpenRead(path))
{
    // ... use stream ...
}
// or for short-lived disposables, just don't dispose:
var stream = File.OpenRead(path);
```

Full notes: `C:\personal_rag\claude_code\lesson_20260506_csx_using_var_declaration_unsupported.md`.

### Script's top-level `var` is null when referenced earlier

**Cause.** `.csx` hoists top-level `var` DECLARATIONS to the top of the
script, but ASSIGNMENTS stay in source order. Referencing a top-level
var before its assignment runs returns `null` at runtime (no compile
error).

**Fix.** Reorder so the assignment runs before any reference. Or use
explicit type declarations to fail at compile time.

Full notes: `C:\personal_rag\claude_code\lesson_20260424_csx_variable_hoisting_trap.md`.

## COM runtime errors

### `COMException 0x800AC472`

**Cause.** Excel-specific "automation operation failed." Common
triggers: Excel is busy with a modal dialog you can't see (compatibility
warnings, AutoSave prompts), or trying to SaveAs while Excel is in an
unfocused/unstable state.

**Fix.** Retry from a fresh PowerShell `New-Object -ComObject` (not the
user-interactive Excel). Confirm no modal dialogs are open. Try
`xlApp.DisplayAlerts = false` before the operation.

### `COMException 0x80010105 RPC_E_SERVERFAULT`

**Cause.** The COM server crashed or rejected the call. Most commonly
in SW automation: calling `IModelDoc2.EditSketch()` on a freshly-pasted
drawing-view sketch.

**Fix.** App-specific. Pattern: attempt once → on COMException, disable
the operation for the rest of the session (circuit breaker) → report
precisely → never retry. Retrying after RPC_E_SERVERFAULT often kills
the app.

Full notes: `C:\personal_rag\solidworks\lesson_20260516_editsketch_rpc_fault_pasted_view_sketch.md`.

### `COMException 0x80010001 RPC_E_CALL_REJECTED`

**Cause.** The COM server (typically Office) is currently busy handling
another call and rejected our incoming one.

**Fix.** Wait + retry. Office apps prefer single-threaded marshaling;
back off 100ms and retry up to a few times.

### `COMException 0x80010105` (re-occurring)

**Cause.** Same as RPC_E_SERVERFAULT. After the first occurrence, the
RPC server is often permanently dead for this process — restart the
target app.

## Performance issues

### `list-sessions` is slow when many docs are open

**Cause.** ROT enumeration walks every entry. With 100+ open Office
documents (across all apps), enumeration time + per-entry bind +
dynamic dispatch can take a second or two.

**Mitigation.** Per-plugin patterns are filtered as we walk. Tighter
regex anchors (`^.*\.xlsx$` instead of `\.xlsx$`) don't help much; the
bind is the expensive part. Consider passing `--session pid:NNNN`
directly when you already know the PID — it skips much of the work.

### Excel COM calls are slow inside a loop

**Cause.** Every property access marshals across the COM boundary. Loops
over individual cells are extremely slow.

**Fix.** Bulk-read with `range.Value` returning `object[,]` (single COM
call). For writes, build a 2D array and assign once to `range.Value`.

```csharp
var arr = new object[1000, 5];
for (int r = 0; r < 1000; r++)
for (int c = 0; c < 5; c++)
    arr[r, c] = $"R{r}C{c}";
xlSheet.Range["A1", "E1000"].Value = arr;   // one COM call
```

## "I don't know which discovery pattern an app uses"

Use the ROT diagnostic in the "no running sessions" section above. Look
at the display names of monikers for the running app:

| Display name format | Pattern | Plugin shape |
|---|---|---|
| File path ending in app's extension | A (document-based) | `RotMonikerPatterns: [@"\.<ext>$"]`, `TryExtractRoot` ascends `.Application` |
| Custom string with PID embedded | B (multi-instance custom) | Anchored regex on the custom format, identity `TryExtractRoot` |
| Nothing visible BUT `[Marshal]::GetActiveObject(ProgID)` works | C (class moniker) | Empty `RotMonikerPatterns`, rely on GetActiveObject fallback |
| ROT empty AND GetActiveObject fails | App doesn't expose COM — abort | — |
