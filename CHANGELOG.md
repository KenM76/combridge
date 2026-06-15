# Changelog

All notable changes to this project will be documented in this file.
Format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/);
versions follow [SemVer](https://semver.org/spec/v2.0.0.html).

## [0.9.0] — `run-script` becomes a clean Unix filter: `ScriptArgs` in, `Stdin` in, stderr separate

Closes TWO FRs that compose into one story:

- `FR_runscript_script_args.md` — argv channel via `ScriptArgs` global
- `FR_runscript_stdin_and_stderr_separation.md` — `Stdin` global +
  stop redirecting `Console.Error` to the output writer

Both FRs had been independently filed but never implemented (the
`Args` FR was misfiled to `Complete\` at some point without code
shipping — verified by repo-wide grep). v0.9.0 ships both together
because they compose: a `.csx` becomes a proper Unix-style
`stdin → stdout(data) + stderr(diag)` filter with argv in and exit
code out (exit-code FR shipped in v0.8.2). The SWBomExcluded
ScripTree provider can now retire its PowerShell wrapper shim and
become plain `combridge solidworks run-script provider.csx -`.

### Added — `IScriptContext` interface

New interface in `ComBridge.Core`:

```csharp
public interface IScriptContext {
    string[] ScriptArgs { get; set; }
    string Stdin { get; set; }
}
```

Each plugin's globals class (`SwGlobals`, `ExcelGlobals`, `WdGlobals`,
`PptGlobals`, `OlGlobals`, plus the four Mac variants) now implements
this interface. The host (`RunScriptCommand`) casts and sets the
fields after `CreateGlobals`. Plugins that DON'T implement it just
skip silently — scripts see empty values, preserving pre-v0.9.0
behavior.

### Added — `ScriptArgs` global (FR 1)

CLI tokens between the script path and the trailing output-file
positional are now available to scripts:

```
combridge solidworks run-script audit.csx --mode quick --offline X: -
                                          └────── ScriptArgs ──────┘
```

Inside the script (both `.csx` and `.vbs`):

```csharp
// .csx
foreach (var arg in ScriptArgs) Console.WriteLine(arg);
return ScriptArgs.Length;
```

```vbscript
' .vbs
For i = 0 To UBound(ScriptArgs)
    WScript.Echo ScriptArgs(i)
Next
```

Empty array when no extra tokens were passed.

### Added — `Stdin` global (FR 2 item 1)

If the calling process redirected stdin to combridge (a pipeline, a
here-doc, a file redirect), the full stream is read eagerly at command
entry and exposed to the script as `Stdin`. Empty string when stdin
isn't redirected.

```csharp
// ScripTree provider .csx — receives request as JSON on stdin:
var req = JsonSerializer.Deserialize<ProviderRequest>(Stdin);
// ... build choices ...
Console.WriteLine(JsonSerializer.Serialize(new { choices, choice_labels }));
```

**Stdin-timeout trap** (caught during smoke-test, important): the
naive implementation `if (Console.IsInputRedirected)
ReadToEndAsync()` **hangs forever** when stdin is inherited-but-empty
(common when combridge is invoked from bash subprocesses, Task
Scheduler, CI runners). `Console.IsInputRedirected` returns `true`
for "non-terminal" — NOT "has data available." The fix:
`ReadStdinWithTimeoutAsync` uses the underlying stream's `ReadAsync`
with a 250 ms cancellation token, extended per chunk so slow
producers aren't truncated. Real producers deliver the first bytes
in microseconds; empty-inherited-handle invocations collapse to "".
See `personal_rag/claude_code/lesson_20260615_console_isinputredirected_inherited_handle_hang.md`.

### Changed — `Console.Error` no longer redirected to `<out>` (FR 2 item 2)

Pre-v0.9.0 `ScriptHost.RunAsync` did:

```csharp
Console.SetOut(output);
Console.SetError(output);   // ← merged stderr into the <out> writer
```

So a script's `Console.Error.WriteLine(...)` corrupted any structured
data on stdout — a provider that needed to emit pure JSON had to
forbid all diagnostics on the success path. v0.9.0 leaves
`Console.Error` alone:

```csharp
Console.SetOut(output);
// Console.Error flows to the process's real stderr (cleaner filter shape)
```

Host-emitted diagnostics (compile errors, `script not found`,
`SCRIPT EXCEPTION`) already go through `output.WriteLine(...)`
directly and are unaffected. Only the SCRIPT's `Console.Error`
changes destination.

### Verified

Full SWBomExcluded provider pattern — stdin in, args in, stdout pure
JSON, stderr diagnostics:

```bash
echo '{"target_file":"foo.SLDDRW"}' | combridge solidworks run-script \
    full_filter.csx --x 1 --y 2 /tmp/out.json 2>/tmp/diag.txt
```

Result:
- `/tmp/out.json` contains clean JSON (parser-safe):
  `{ "received_target": "foo.SLDDRW", "arg_count": 4, "args": ["--x","1","--y","2"] }`
- `/tmp/diag.txt` contains the diagnostic:
  `[diag] received 29 bytes on stdin, 4 args`
- combridge exits 0

Plus the empty-stdin no-hang case (Test 1), real-pipe case (Test 2),
and stderr-split case (Test 3) — all pass.

### Impact on existing scripts

Mostly none — both fields default to empty when unused. The one
behavior change worth noting: any existing `.csx` that wrote
diagnostics via `Console.Error.WriteLine(...)` expecting them to land
in `<out>` will now write to the process's real stderr instead. This
is the Unix-correct behavior; scripts that need the old merge can
explicitly redirect with `Console.SetError(Console.Out)` at the top
of the file.

### Files

Added:
- `src/ComBridge.Core/IScriptContext.cs`

Changed:
- `src/ComBridge.Core/Commands/RunScriptCommand.cs` — populates
  `ScriptArgs` + `Stdin` on globals before invoking host;
  `ReadStdinWithTimeoutAsync` helper
- `src/ComBridge.Core/ScriptHost.cs` — removed `Console.SetError`
- `src/plugins/ComBridge.Plugins.SolidWorks/SolidWorksPlugin.cs` —
  `SwGlobals` implements `IScriptContext`
- `src/plugins/ComBridge.Plugins.Excel/ExcelPlugin.cs` —
  `ExcelGlobals` implements `IScriptContext`
- `src/plugins/ComBridge.Plugins.Word/WordPlugin.cs` —
  `WdGlobals` implements `IScriptContext`
- `src/plugins/ComBridge.Plugins.PowerPoint/PowerPointPlugin.cs` —
  `PptGlobals` implements `IScriptContext`
- `src/plugins/ComBridge.Plugins.Outlook/OutlookPlugin.cs` —
  `OlGlobals` implements `IScriptContext`
- `src/plugins/ComBridge.Plugins.Excel.Mac/XlMacApp.cs` —
  `XlMacGlobals` implements `IScriptContext`
- `src/plugins/ComBridge.Plugins.Word.Mac/WdMacApp.cs` —
  `WdMacGlobals` implements `IScriptContext`
- `src/plugins/ComBridge.Plugins.PowerPoint.Mac/PptMacApp.cs` —
  `PptMacGlobals` implements `IScriptContext`
- `src/plugins/ComBridge.Plugins.Outlook.Mac/OlMacApp.cs` —
  `OlMacGlobals` implements `IScriptContext`

Both FRs moved to `Complete\` with implementation log stamps.
The `Args` FR's old "Status: PROPOSED" header was preserved with a
note that it had been misfiled prior to actual implementation.

## [0.8.2] — `run-script` now propagates the script's `return N` as the process exit code (contract fix)

Closes `FR_runscript_propagate_script_return_value.md`. A `.csx`
written as `Console.WriteLine("probe"); return 5;` now exits with
code **5**, not 0. The documented behavior in `ScriptHost.RunAsync`'s
XML remarks (and in `LLM/scripting.md` § "Exit codes from scripts" —
"Returned `int` becomes the script-host's exit code") was a contract
the implementation hadn't actually been honoring: every successful
script run was returning 0 regardless of its return value.

### Impact

ScripTree drivers and shell callers can now key red/green status off
combridge's exit code without parsing the script's output text for
status markers. The FR was filed from the MergeInstanceMates ScripTree
tool — `run_merge_instance_mates.py` was parsing stdout for `ERROR:`
/ `FAIL` strings to reconstruct the status that should have come
through `$?`. Every future status-bearing `.csx` would have needed the
same workaround until this shipped.

### Precedence

Bridge-level reserved codes (2 file-not-found, 3 compile error,
4 script exception, 5 host failure) short-circuit before the script's
return value is read. The "return N" path is reached ONLY when the
script ran to completion. Reserved-code collisions (e.g. a script
returning 4) are the author's problem to avoid — the documented rule
"non-zero = failure" advises against reusing them for script signaling.

### Verified

- `return 5;` → exit 5 ✓
- `return 42;` → exit 42 ✓
- no `return` statement → exit 0 ✓
- `throw` before `return 99;` → exit 5 (HOST EXCEPTION precedence; the unreached return is correctly ignored) ✓

### VBScript path unaffected

The `.vbs` host already propagates `WScript.Quit(N)` as the exit code
(verified in v0.8.0's smoke test: `WScript.Quit 42 → exit 42`). This
patch only touches the Roslyn `.csx` path.

## [0.8.1] — cement null-global skip with empirical crash finding + actionable warning

Implements `FR_vbscript_scripting_host.md`. Adds a second script engine
to `run-script` so `.vbs` files run against the same plugin globals
(`swApp`/`swDoc`/`xlApp`/etc.) the Roslyn `.csx` host exposes. Same
`--session` attach, same output capture, same exit-code mapping. The
`.csx` path is unchanged.

### Why this exists

SolidWorks's entire automation ecosystem is VBA. Every forum macro,
every recorded macro, every shop's library, every GoEngineer/TriMech
tutorial. Forcing C# rewrites blocks all of that from joining the
combridge ecosystem — and the rewrite is expensive (per-API: interop
interface name, every enum's integer value, the `out`-vs-`ref` PIA
quirks documented across `personal_rag/solidworks/`). VBScript late
binding sidesteps every one of those quirks because COM `ByRef`
out-params match the COM ABI natively. For SW automation specifically,
**VBScript is more robust than typed C# interop, not less**.

### Architecture

`ScriptHost.RunAsync` now dispatches by file extension:

| Extension | Engine |
|---|---|
| `.csx` | Roslyn C# (existing, unchanged) |
| `.vbs` | New IActiveScript-hosted VBScript engine |
| `.vba` / `.bas` / `.swp` | Rejected with a clear error (VBA is NOT VBScript; UserForms / `Type` / `Public Const` etc. won't parse — convert to `.vbs` or rewrite as `.csx`) |
| anything else | Rejected with "supported: .csx, .vbs" |

The VBScript engine is hosted via `CoCreateInstance(CLSID_VBScript)` →
`IActiveScript` + `IActiveScriptParse64`. The host site
(`VbScriptSite`) reflects the plugin's globals object and registers
each public reference-type instance property as a named script item
via `AddNamedItem` / `GetItemInfo`. No msscript.ocx dependency, runs
in-process at 64-bit.

### Scope changes vs. the FR (deliberate)

The FR proposed several conveniences I deliberately deferred to keep
v0.8.0 focused on the core engine. Documented in the FR's
implementation log:

- **No `.vba` / `.bas` / `.swp` extension aliasing.** Those are VBA
  file extensions, not VBScript. Accepting them by extension would
  produce confusing parse errors on syntax the VBScript engine doesn't
  recognize (UserForms, `Type` declarations, `Public Const`, sigil-
  prefixed sub visibility). Rejected at the extension dispatcher with
  a one-paragraph message pointing the author at the conversion path.
- **No `WScript.Arguments` collection.** `ScriptArgs` (already shared
  with `.csx`) is the supported way to pass argv. If a real cscript-
  authored macro needs it later, easy v0.8.1 addition.
- **No `swconst` pre-injection.** Use integer literals (standard SW-
  VBScript practice; matches every existing standalone `.vbs`).

### Exit-code mapping (parallel to `.csx`)

| Code | Meaning |
|---|---|
| `0` | script ran to completion (or called `WScript.Quit 0`) |
| `2` | script file not found |
| `3` | VBScript syntax/parse error |
| `4` | runtime error (`Err.Raise`, division by zero, unbound name, COM exception) |
| `5` | host failure (engine `CoCreateInstance` failed, e.g. VBScript removed from this Windows) |
| any other `int` | value passed to `WScript.Quit(N)` |

Phase detection (parse vs runtime) uses `OnEnterScript` — NOT
`OnStateChange(SCRIPTSTATE_CONNECTED)`. The CONNECTED state-change
fires when the script COMPLETES, not when it starts; using it for
phase classification mis-labels every runtime error as a parse error.
Caught by a live test (division-by-zero showed exit 0 with "PARSE
ERROR" tag) and fixed before v0.8.0 shipped.

### Globals injection limitations

`IActiveScriptSite::GetItemInfo` returns named items as IUnknown for
IDispatch wrapping. Two property categories can't be returned that way
and are skipped with a host-side warning written above script output:

- **Value-type properties** (boxed primitives, enums): can't be
  returned as IUnknown. `SwGlobals.swDocType` (a `swDocumentTypes_e`
  enum) hits this. Scripts that need the value read it via the typed
  COM wrapper (e.g. `swDoc.GetType`).
- **Null reference-type properties**: returning IUnknown=null produces
  an "unknown name" runtime error, not a Nothing-equivalent.
  `swPart`/`swAssy`/`swDrawing` are skipped when no doc of that type is
  active. Scripts can guard via `If swDoc Is Nothing Then ...`.

Both skip lists are printed at run start so the author isn't mystified
when a name they expect comes back as undefined.

### VBScript deprecation — honest disclosure

Microsoft formally deprecated VBScript in 2024 with planned removal
from a future Windows release. This host depends on the in-box
`vbscript.dll`. When Microsoft removes it, `CoCreateInstance` will
return `REGDB_E_CLASSNOTREG` and this command will exit 5 with a
message pointing to `.csx`. Building on a deprecated runtime is the
right call here (the existing VBA macro corpus is too valuable to
leave unintegrated while we still can), but consumers should know
they're investing in a runtime with a known sunset.

### Verified

Live-tested against the running SolidWorks 2026 SP1.1 session:

```vbscript
WScript.Echo "SolidWorks version: " & swApp.RevisionNumber
If swDoc Is Nothing Then WScript.Echo "no doc": WScript.Quit 0
WScript.Echo "Active doc title: " & swDoc.GetTitle
WScript.Echo "Active doc path: " & swDoc.GetPathName
```

Output:
```
# vbscript host: value-type globals not injected (read via typed wrapper): swDocType
# vbscript host: null globals not injected (no active doc/etc.): swPart, swAssy
SolidWorks version: 34.1.1
Visible: True
Active doc title: TS-0220-192192 - TS-0220-192192
Active doc path: W:\Engineering\Products\TS-0220 Zacon BVD Door in Door - Top Level Assemblies\TS-0220-192192\TS-0220-192192.SLDDRW
```

Parse-error, runtime-error, and `WScript.Quit(N)` exit-code paths all
verified.

### Architecture future-proofing

The COM interop layer (`ActiveScriptInterop.cs`) declares CLSIDs for
BOTH VBScript and JScript (the latter unused in v0.8.0). Adding a
JScript host in the future is one extension-dispatch case +
`Type.GetTypeFromCLSID(CLSID_JScript)` instead of `CLSID_VBScript` —
the rest of the hosting code (site, parser, engine driver) is
language-agnostic.

PowerShell would be a different hosting story
(`System.Management.Automation` Runspace, not `IActiveScript`) — left
for a future FR. Python would be different again (pythonnet or
subprocess) — also future FR.

### Files added
- `src/ComBridge.Core/ActiveScriptInterop.cs` — COM interop layer
  (`IActiveScript`, `IActiveScriptParse64`, `IActiveScriptSite`,
  `IActiveScriptError`, `EXCEPINFO`, CLSIDs, constants)
- `src/ComBridge.Core/VbScriptEngine.cs` — engine driver +
  `VbScriptSite` (IActiveScriptSite implementation) + `WScriptShim`
  (`Echo` / `Quit`)

### Files changed
- `src/ComBridge.Core/ScriptHost.cs` — extension-dispatcher prepended
  to `RunAsync`. Roslyn path unchanged below the switch.

## [0.7.0] — Outlook search v2: multi-term + word matching + sender + EntryID, and new `outlook get`

Addresses `FR_outlook_search_v2_multiterm_sender_match_and_get.md` in full
(all 6 items). Both Windows and macOS Outlook plugins updated for parity.

**Breaking change** in CLI output (justified per the FR's "we're the only
users" note — backward compatibility deliberately not preserved). The
`outlook search` columns now include `matched`, `entryid`, and `storeid`
on every row (no longer flag-gated), and defaults changed where the old
default was wrong. Wrappers built against v0.6.x output need their
column expectations updated.

### Why withdraw the v0.4.0 behavior

The v0.4.0 implementation was single-term, substring-only, subject/body-only,
since-only, and emitted no EntryID. The FR documented a real
"cast a wide net" task (searching for a Gasspring.ca order across two
mailboxes for a US$313.89 charge) where every one of those gaps blocked
progress. Specifically, substring matching of `pdac` against a base64
URL-tracking blob `…ZPDACfM5…` matched a Sudbury meal newsletter — the
kind of systematic false positive modern marketing mail generates by
design.

### Changed defaults (FR § "What I'd ship instead")

| Setting | Old default | New default | Why |
|---|---|---|---|
| `--match` | implicit substring | **`word`** (ci_phrasematch on indexed; LIKE+regex fallback) | substring is just wrong for marketing mail; the fix should be the default |
| `--fields` | `subject,body` | **`subject,body,from`** | if you can search the sender for free, you almost always want to |
| EntryID/StoreID emission | not emitted | **always emitted** | connects `search` → `get` without a flag |

### Added — `outlook search` v2

- **Multi-term**: `--query` is now repeatable AND accepts comma-separated terms.
  Both forms compose. `--query "a,b" --query c` → three terms ORed.
- **Sender field**: `--fields` accepts `from` (alias `sender`), mapping to
  both `urn:schemas:httpmail:fromname` (display) and `fromemail` (SMTP
  address). Default fields now include `from`.
- **Word matching**: `--match word|substring`. `word` (default) uses DASL
  `ci_phrasematch` on content-indexed stores. Per-folder try/catch falls
  back to `LIKE` + C#-side `\bterm\b` regex when `ci_phrasematch` throws
  "condition is not valid" (the non-indexed-store signal). Either path
  produces the same word-boundary semantics.
- **Date window**: `--until yyyy-MM-dd` mirrors `--since`. Together they
  form a closed interval. Either alone is open on the other end.
- **`matched` column**: every row lists the term(s) that hit, computed
  by a C#-side re-scan of each Restrict candidate against
  (term × field) regex pairs. The same re-scan drops word-mode false
  positives that DASL LIKE let through but `\bterm\b` rejects.
- **`entryid` + `storeid` columns**: always emitted on every row. EntryIDs
  are only unique within a store, so both are needed to call
  `GetItemFromID(entryID, storeID)` reliably.

### Added — new `outlook get` command

Resolves an item by either path:

- `--id <EntryID> [--store <substr>]` — direct fetch via
  `NameSpace.GetItemFromID(entryID, storeID)`. Fast and precise. The
  `--store` substring is resolved to the matching store's `StoreID`
  before lookup; without it, GetItemFromID searches the default store
  and may miss items in other accounts.
- `--subject <substr> [--store <substr>] [--folder <substr>] [--max N]` —
  recursive walk for interactive use when you don't have an EntryID
  handy. Returns the first `--max` matches (default 1).

Output is a one-block-per-item plain-text dump with labeled headers
(`Subject`, `From`, `To`, `Cc`, `ReceivedTime`, `Size`, `Folder`,
`EntryID`, `StoreID`) followed by `[BODY]` with the plain-text content.
`--html` additionally dumps `[HTMLBODY]`. `--headers` adds the
attachments list (name + size) and `MessageClass`.

### Verified (live test from the FR's driving task)

Ran the exact Gasspring scenario from the FR against the live mailboxes:

```
combridge outlook search --query "ace control,acecontrols,313.89,gasspring,forklift,spring" \
    --match word --since 2026-02-01 --until 2026-03-31 --snippet  /tmp/gasspring.tsv
```

- **14 hits emitted** (vs the FR's documented 4,412-hit substring flood — a
  99.7% noise reduction from the word-boundary filter)
- **First hit** = `Your Gasspring.ca order WS14080CA has been received!` with
  `matched=313.89,gasspring,spring` (three signal terms — exactly the FR's
  predicted top-row pattern)
- **4 stores walked, 257 folders walked, 0 folders required fallback**
  (entire mailbox content-indexed)
- **Per-hit EntryID + StoreID emitted in every row**

Then `outlook get --id <EntryID> --store toprops --headers` fetched the
full order email including:
- Headers (Subject, From, To, ReceivedTime, Size, Folder, EntryID, StoreID, MessageClass)
- Attachment list (`MountingDrawing_WS14080CA_8-19-160_8880.pdf`, 200 KB)
- Full plain-text body with the line items the FR predicted
  (4× $153.48 gas springs, 8× $59.36 brackets, **Total: US$313.89**)

Both subject-fallback (`--subject "WS14080CA" --store toprops`) and
direct-ID paths return the same item.

### Mac parity

`OlMacSearchCommand` rewritten to match. Same flag surface, same column
shape, same `matched`/`entryid`/`storeid` always-emitted output. Mac
differences (architectural, not behavioral):

- AppleScript `whose` is substring-only — no `ci_phrasematch`. Word-mode
  always uses the C#-side `\bterm\b` regex post-filter. There's no
  "fast path" for indexed stores; every search behaves like the Windows
  fallback path.
- Per-term separate `whose` queries because Mac Outlook's dictionary
  doesn't reliably accept compound `whose ... or ...` predicates against
  messages. Costs `terms × fields` AppleScript evaluations per folder
  instead of one — measurable on large mailboxes.
- The `entryid` column is the integer AppleScript exposes via
  `id of message` (not the opaque hex EntryID Windows uses). Mac has no
  StoreID concept; we reuse the account name in the `storeid` column for
  cross-OS schema compatibility.
- `OlMacGetCommand` (new) accepts `--id <integer>` and resolves via
  AppleScript `messages whose id is N` walked per account/folder.
  `--subject` path identical to Windows.
- Still classic Outlook for Mac only — "New Outlook for Mac" (the 2024+
  Catalyst UI) restricts the AppleScript dictionary too far.

### Removed
- v0.4.1's `OlMacApp.Search(...)` programmatic helper and `SearchHit`
  record. The new search command inlines its own AppleScript; a
  programmatic equivalent should either shell out to
  `combridge outlook search` or write its own AppleScript via
  `Osascript.Run(...)`. Less mechanism, fewer abstractions.

### Files
- `src/plugins/ComBridge.Plugins.Outlook/OlSearchCommand.cs` (rewritten,
  new file split from OutlookPlugin.cs)
- `src/plugins/ComBridge.Plugins.Outlook/OlGetCommand.cs` (new)
- `src/plugins/ComBridge.Plugins.Outlook/OutlookPlugin.cs` (old inline
  OlSearchCommand removed; Commands list updated)
- `src/plugins/ComBridge.Plugins.Outlook.Mac/OlMacSearchCommand.cs` (rewritten)
- `src/plugins/ComBridge.Plugins.Outlook.Mac/OlMacGetCommand.cs` (new)
- `src/plugins/ComBridge.Plugins.Outlook.Mac/OutlookMacPlugin.cs` (old
  inline OlMacSearchCommand removed; Commands list updated)
- `src/plugins/ComBridge.Plugins.Outlook.Mac/OlMacApp.cs` (Search +
  SearchHit removed)

## [0.6.0] — `list-addins` across every plugin (universal diagnostic command)

Adds a new `list-addins` subcommand to every plugin: Excel, Word,
PowerPoint, Outlook, SolidWorks (Windows) plus best-effort Excel.Mac and
Word.Mac via AppleScript. Same diagnostic category as `list-sessions` /
`info` — universal infrastructure, machine-parsable TSV output, no
business logic baked in.

### Why this exists

"Is the Toolbox add-in actually loaded?" / "What COM add-ins is this
Excel instance running?" / "Did Acrobat PDFMaker install correctly?"
These are recurring diagnostic questions every consumer of an Office or
SolidWorks plugin asks. Today each consumer has to write 10-30 lines of
COM/registry enumeration in a `.csx` to answer them — each app has its
own non-obvious enumeration model (COMAddIns + AddIns split in Office,
dual registry tree + per-version hidden tree in SolidWorks, etc.).
Shipping it as a built-in turns N rediscoveries into one canonical
answer with a stable output shape.

### Output shape (TSV, consistent across plugins)

```
# columns: name<TAB>id<TAB>loaded<TAB>kind<TAB>description
<name>\t<id>\t<true|false>\t<COM|XLL|VBA|WLL|TEMPLATE|NATIVE>\t<extra>
...
# total: <N>
```

Header rows are prefixed with `#` for easy grep/awk filtering. Tabs/
newlines inside any field are replaced with spaces so consumers can
split on `\t` without escape handling.

### Per-plugin coverage

| Plugin | Enumeration source | Notes |
|---|---|---|
| `excel` | `Application.COMAddIns` + `Application.AddIns` | Both COM/VSTO and XLL/.xla/.xlam in one stream. Each collection wrapped in its own try/catch so a partial failure (security policy denying one collection) still emits the other. |
| `word` | `Application.COMAddIns` + `Application.AddIns` | COM + global templates (.dot/.dotm) + WLLs. |
| `powerpoint` | `Application.COMAddIns` + `Application.AddIns` | COM + .ppam/.ppa. Uses `MsoTriState` for the loaded flag (PowerPoint distinguishes registered-but-not-loaded from loaded-this-session). |
| `outlook` | `Application.COMAddIns` only | No equivalent classic-addin collection. Newer security-hardened deployments may restrict access — emits a WARN row rather than failing. |
| `solidworks` | Registry walk + `ISldWorks.GetAddInObject(Clsid)` | See below — substantially more complex than the Office model. |
| `excel.mac` | AppleScript `addins of application` | Best-effort, strict subset of Windows (no COMAddIns, no XLL on macOS). Same TSV column shape so cross-OS ScripTree apps work. |
| `word.mac` | AppleScript `add-ins of application` | Same as Excel.Mac: subset of Windows. |
| `powerpoint.mac` / `outlook.mac` | — | Not shipped. PowerPoint/Outlook for Mac's AppleScript dictionaries don't expose addin collections. Could be added if a use case appears. |

### SolidWorks `list-addins` — the interesting one

SW has no first-class `GetAddIns()` API. The canonical answer is a
dual registry walk plus per-add-in probes, fully documented in
`C:\personal_rag\solidworks\lesson_20260529_sw_addin_dual_registry.md`.
The implementation honors every finding from that lesson:

- **Two HKLM trees walked**: `HKLM\SOFTWARE\SolidWorks\AddIns\{guid}`
  (UI-visible, the ones Tools → Add-Ins shows) AND
  `HKLM\SOFTWARE\SolidWorks\SOLIDWORKS <ver>\Addins\{guid}` (hidden
  product-feature add-ins: Design Checker, Costing, TolAnalyst,
  Sustainability, Reveng/ScanTo3D, etc.). We auto-discover every
  installed SW version's hidden tree so multi-version installs surface
  every version's set with version-tagged scope.
- **Per-user enabled-at-startup state** lives at
  `HKCU\Software\SolidWorks\AddInsStartup\{guid}\(Default)` REG_DWORD
  (NOT under the AddIns key — common mistake an earlier draft of this
  code made). Missing key = effectively disabled.
- **Currently-loaded probe** via `ISldWorks.GetAddInObject(Clsid)`. The
  canonical RAG (`sldworks_methods_v3_llm.rag:82`) confirms the
  parameter is the CLSID, not the ProgID — a different gotcha the
  first draft hit. Returns the live IDispatch when loaded, null
  otherwise. Tolerated per-add-in try/catch for 3rd-party addins that
  throw on the probe.
- **DLL path resolution** handles the .NET-hosted `mscoree.dll` →
  `CodeBase` URL indirection automatically. Most modern SW add-ins
  (and all 3rd-party managed ones) hit this path; the column shows
  the actual assembly path, not "mscoree.dll".
- **Friendly name** comes from `HKCR\CLSID\{guid}\(Default)` (the
  COM-registered class name like "SWDesignCheck Class"), NOT from the
  HKLM AddIns subkey. Same source the SW Tools → Add-Ins UI reads.
- **Case-insensitive CLSID dedup** across hives — registry GUID casing
  varies between HKLM and HKCU.

Live-verified on the dev machine: enumerated 25 addins total — 9
UI-visible (Composer, OpenToolbox.Addin, SwClaudeAddinPro, 3DExpExchange,
etc.), 11 hidden:2026 matching the exact set listed in the lesson
(Autotrace/Picture2Sketch, Aura, AutoDrawings, CircuitWorks, Costing/
SwcAddin, PartReviewer, Sustainability/swgApp, Design Checker/
SWDesignCheck, TolAnalyst, Reveng/ScanTo3D, sldxps), plus 5 hkcu-only
including FuncFeatApp lazy-loaded into the live session (matched the
lesson's prediction about MacroFeature-triggered lazy loading).

Not enumerable: the ~4 modules hardcoded into `SLDWORKS.EXE` itself
(`fworks.dll`/FeatureWorks, `swbrowser.dll`/Toolbox Browser, etc.).
They appear in neither registry tree and can't be toggled. The output's
trailing `# total:` line notes this explicitly so consumers don't think
their machine is missing entries.

### Files added
- `src/plugins/ComBridge.Plugins.Excel/ListAddinsCommand.cs`
- `src/plugins/ComBridge.Plugins.Word/ListAddinsCommand.cs`
- `src/plugins/ComBridge.Plugins.PowerPoint/ListAddinsCommand.cs`
- `src/plugins/ComBridge.Plugins.Outlook/ListAddinsCommand.cs`
- `src/plugins/ComBridge.Plugins.SolidWorks/ListAddinsCommand.cs`
- `src/plugins/ComBridge.Plugins.Excel.Mac/ListAddinsCommand.cs`
- `src/plugins/ComBridge.Plugins.Word.Mac/ListAddinsCommand.cs`

Each plugin's main `*.Plugin.cs` adds the new command to its `Commands`
collection. No contract changes; no breaking changes.

### Verified
- `excel list-addins` against live Excel: 8 addins enumerated (PowerMap,
  Power Pivot, Acrobat PDFMaker, Data Streamer, plus XLL/VBA Analysis
  ToolPak / Solver / Euro Tools). Loaded vs not-loaded correctly reflected.
- `solidworks list-addins` against live SW 2026 SP1.1: 25 addins total
  matching the dual-registry lesson's predictions.
- All seven plugins build clean; the four Windows Office plugins
  compile against `Microsoft.Office.Core.COMAddIn` (shared Office
  plumbing, not per-app namespace — a build error caught the wrong
  assumption mid-implementation).
- `list-commands` shows `list-addins` as `(plugin)` source for every
  plugin that has it.

## [0.5.0] — withdraw v0.4.2 alias preamble; ship visible scaffolding + smart CS0104 hints

**Breaking change** in the plugin contract (justifies the minor-version
bump): `IComBridgePlugin.ScriptUsingAliases` is REMOVED. Plugins built
against v0.4.2's contract that relied on the preamble mechanism need to
be rebuilt against v0.5.0; the four Office plugins shipped in this repo
are already updated.

### Why withdraw v0.4.2

The v0.4.2 mechanism injected `using Xl = global::Microsoft.Office.Interop.Excel;`
into the script source before Roslyn compiled it. That solved the CS0104
papercut, but at app-store scale (thousands of plugins, thousands of
authors, public script catalog) it broke the source-is-truth contract
in ways that compound:

- **External IDEs can't see the preamble.** Devs editing .csx files in
  VS Code, Rider, or Cursor saw red squiggles under `Xl` even though
  the script ran fine. No IDE knew about combridge's host injection.
- **LLMs reading the .csx in isolation hallucinated.** They saw `Xl.Range`
  with no `using Xl = ...` line and tried to "fix" what wasn't broken.
- **App-store auditors couldn't evaluate published scripts.** Reading
  the source required knowing host internals. That's a market-friction
  tax on every install decision.
- **Roslyn-format fragility.** The line-number remap regex depended on
  Roslyn's diagnostic format staying stable; a future Roslyn version
  could silently misreport error locations.
- **Mechanism attracts mechanism.** "Rewrite the script before compile"
  is a feature surface that grows. We don't want it.

The doc-fix-alone equivalent left a one-line-per-script onboarding cost.
v0.5.0 eliminates that cost with two visible tools that don't break the
source-is-truth contract.

### Added — visible scaffolding (replaces the preamble)

- **`<plugin> new-script <path> [--force]`** — every Windows Office
  plugin (Excel, Word, PowerPoint, Outlook) now ships a `new-script`
  subcommand that scaffolds a starter `.csx` with the alias line, a
  header comment documenting the available globals, and a minimal
  example body. Edit the body, run it. The alias declaration lives in
  the file the author owns — every reader (IDE, LLM, auditor, future
  maintainer) sees exactly what's in scope.
- **`ScriptScaffold.WriteTemplate`** in `ComBridge.Core` — shared
  helper that all `new-script` commands delegate to. Parses `<path>
  [--force]`, refuses to overwrite without `--force`, writes the
  template, reports the result with a "now run it with:" hint. Future
  plugins (Visio, AutoCAD, Inventor, etc.) get scaffolding by
  delegating one line and supplying their template constant.

### Added — smart CS0104 hints (replaces the line-number remap)

- **`AugmentOfficeDiagnostic`** in `ScriptHost` — when Roslyn produces
  a CS0104 ambiguous-reference error for an Office-interop / BCL
  collision, the host detects the pattern and appends a one-line hint
  with the exact `using` to add and the qualified form to use:

  ```
  collision_test.csx(2,1): error CS0104: 'Range' is an ambiguous reference between 'Microsoft.Office.Interop.Word.Range' and 'System.Range'
    -> Hint: add this to the top of your script:
           using Wd = global::Microsoft.Office.Interop.Word;
         then use 'Wd.Range' instead of bare 'Range',
         or qualify the BCL side as 'System.Range'.
         See LLM/scripting.md for the full collision table.
  ```

  Purely additive — the original diagnostic, including its
  `(line,col)` span, is preserved verbatim. No rewriting, no
  remapping.

### Removed

- `IComBridgePlugin.ScriptUsingAliases` contract member
- `ScriptHost`'s preamble injection (`using Xl = ...; ` prepended to
  script source)
- `ScriptHost.DetectEncoding` (no longer needed — Roslyn's
  `File.OpenRead` overload handles encoding)
- `ScriptHost.RemapDiagnosticLine` + `DiagLocRx` (no preamble means
  Roslyn's reported line numbers already match the author's source)
- `ScriptUsingAliases` overrides on the four Windows Office plugins

### Verified

- All four Office plugins' `new-script` writes a valid starter that
  COMPILES AND RUNS on its own. Excel scaffold ran live against an
  open workbook and printed sheet stats.
- Refuse-to-overwrite (exit code 1) works; `--force` overwrites
  cleanly; missing-directory case (exit code 2) reports clearly.
- CS0104 hint augmentation: a Word .csx with bare `Range r;` now
  reports the original error PLUS the actionable hint, with the
  author's actual line `(2,1)` preserved (no remap needed).
- Non-CS0104 errors pass through unchanged — only Office-interop /
  BCL collisions are augmented.

### Docs
- `LLM/scripting.md` § Office namespace shadowing rewritten:
  leads with the explicit `using Xl = ...` pattern (the recommended
  convention), documents `new-script` as the zero-typing workflow,
  describes the CS0104 hint as the recovery path. Includes a candid
  paragraph on why v0.4.2 was withdrawn — the source-is-truth contract
  matters more than saving the author one line.
- `LLM/troubleshooting.md` CS0104 entry updated to lead with the
  alias-declaration fix and the `new-script` shortcut. Historical note
  references the rejected FR.
- `FR_office_script_interop_alias.md` moved to
  `Rejected/` with the rejection rationale stamped on top.
- `FR_scripting_dx_and_outlook_search.md` moved to `Complete/`.

## [0.4.2] — auto-provided interop aliases for Office scripts (WITHDRAWN in v0.5.0)

**This release was withdrawn.** See v0.5.0 above for the reasoning and
the visible-scaffolding replacement. The release tag/notes remain on
GitHub for history; do not depend on the API surface that shipped here.



Addresses `D:\Dev\FeatureRequests\ComBridge_FeatureRequests\FR_office_script_interop_alias.md`
in full. Implements the FR's primary proposal (option **b** —
plugin-contributed aliases rendered into a host preamble) rather than the
doc-only fallback.

### Why this exists

Every Windows Office `.csx` hit the same CS0104 pain point: the interop
namespace defines its own `Range`, `Exception`, `Application`, `Style`,
`Font`, `Action`, `Page`, etc., colliding with the same-named BCL types.
`Range` is the worst offender because modern C# added `System.Range` for
slicing syntax — so the single most common Office idiom
(`Range used = xlSheet.UsedRange;`) failed to compile until every
script author re-typed `using Xl = global::Microsoft.Office.Interop.Excel;`
at the top. The FR identified this pattern sitting latent in 15 scripts
across the ScripTreeApps catalog, surfacing only on first run.

### Added
- **`IComBridgePlugin.ScriptUsingAliases`** — new optional contract
  member. Returns alias bodies (e.g. `"Xl = global::Microsoft.Office.Interop.Excel"`);
  the host renders them as `using <alias>;` directives. Default = empty,
  so existing plugins (SolidWorks, all Mac plugins) compile and run
  unchanged with zero behavior change.
- **Alias preamble** in `ScriptHost.RunAsync` — concatenates each
  plugin's contributed aliases onto a single first line of the script
  source, preserving BOM + encoding so PDB emit still works
  (CS8055-free) and non-ASCII characters in script bodies round-trip
  intact. Roslyn `ScriptOptions.Imports` accepts namespaces but not
  alias directives, so the preamble route is the only working option.
- **Diagnostic line-number remapping** — `RemapDiagnosticLine` rewrites
  the `(LINE,COL)` span in Roslyn diagnostic strings so compile errors
  point at the author's real source. A CS0104 reported at compiled
  line 168 surfaces as line 167 (matching what the author sees in their
  editor). Errors inside the preamble itself (a plugin-author bug,
  not a script-author bug) are left untouched so they're loud.
- **Four Windows Office plugins now contribute their alias** —
  `excel` → `Xl`, `word` → `Wd`, `powerpoint` → `Pp`, `outlook` → `Ol`
  (all qualified to `global::Microsoft.Office.Interop.*`). Mac plugins
  contribute nothing — they don't have an interop namespace to shadow.

### What stays the same (deliberately)

- **Bare `Range` is still ambiguous.** The alias only guarantees a
  reliable qualifier (`Xl.Range`, `Wd.Range`) is always in scope; it
  does NOT silently pick the Office type over `System.Range`. This is
  the FR's explicit acceptance criterion — we don't want to win the
  bare-name race for the author.
- **`System.Exception` etc. still need to be qualified** (or you can
  catch a non-colliding subtype like `COMException`). The new aliases
  fix the Office-side qualifier, not the BCL side.
- **Existing scripts that already declare `using Xl = …` keep working.**
  Roslyn quietly accepts the duplicate (same alias to the same
  namespace); we don't conflict with manual declarations.

### Verified

- Positive: a Word `.csx` containing `typeof(Wd.Application).FullName`
  compiles and runs with NO author-declared alias.
- Negative: a Word `.csx` containing bare `Range r;` still fails with
  CS0104 between `Microsoft.Office.Interop.Word.Range` and
  `System.Range` — confirming we didn't accidentally resolve the race.
- Line-number remap: the CS0104 from the negative test surfaces at
  `(2,1)` (the author's actual line in the .csx), not `(3,1)` —
  proving the preamble offset is being subtracted.

### Docs
- `LLM/scripting.md` § "Office interop namespaces shadow common BCL
  names" rewritten to lead with the auto-provided alias (recommended)
  and show the fully-qualify fallback. New alias-mechanics note
  explains the preamble + line remap for plugin authors.
- `LLM/troubleshooting.md` gained a dedicated CS0104 entry covering
  Office-interop / BCL collisions with the alias table inline and the
  rationale for keeping bare `Range` ambiguous.
- `FR_office_script_interop_alias.md` stamped with implementation log
  noting which option was taken (b, not a) and why.

## [0.4.1] — outlook search on macOS

Lifts the v0.4.0 deferral. The Mac Outlook plugin now ships a `search`
command with the same flag surface as the Windows version, so a single
ScripTree `.scriptree` wrapping `combridge outlook search ...` works on
both OSes.

### Added
- **`outlook search` command** on `ComBridge.Plugins.Outlook.Mac` —
  AppleScript-driven recursive mail search. Same flags as the Windows
  Outlook plugin: `--query`, `--store`, `--folder`, `--fields`, `--max`,
  `--since`, `--snippet`. Same TSV output columns
  (`date, account, folder, sender, subject, [snippets]`) so downstream
  parsing is OS-agnostic.
- **`OlMacApp.Search(...)`** — programmatic Mac search API for `.csx`
  scripts. Returns `List<SearchHit>` records (date, account, folder
  path, sender name, sender address, subject, body). Body is only
  fetched when `wantBody: true` since AppleScript body access is slow.

### Implementation details
- One big `osascript` invocation walks every Exchange/IMAP/POP account's
  full folder tree (recursive AppleScript handler) and returns a
  delimited blob using `␞` (U+241E) field separator + `␝` (U+241D) row
  separator — characters extremely unlikely to appear in mail bodies.
- The "subject vs body" filter is split into TWO separate `whose`
  passes per folder (subject-contains, then content-contains) because
  Outlook for Mac's AppleScript doesn't reliably accept compound
  `whose ... or ...` predicates against `messages`.
- Date filter is post-fetched in the script (AppleScript `whose` on
  `date` comparisons against `time received` is locale-finicky); the
  C# code formats the `--since` argument as `MM/DD/YYYY HH:MM`.
- Snippet extraction (when `--snippet` is passed) is identical to the
  Windows version: collapse whitespace → `Regex.Matches` → ±60-char
  windows around each hit → cap at 3 non-overlapping windows per
  message. Body is fetched per-hit which is the slow path; skip
  `--snippet` if you only need the headers.

### Performance vs. Windows
- AppleScript `whose` is server-side for Exchange (acceptable) but
  client-side for IMAP/POP (significantly slower than DASL `Restrict`).
- Expect a query that finishes in ~50 ms on Windows DASL to take
  several seconds on Mac AppleScript against the same mailbox.
- Mitigation: always scope down with `--store`, `--folder`, and
  `--since` for interactive use.

### Caveats (Mac-only)
- Targets **classic Outlook for Mac only**. "New Outlook for Mac" (the
  2024+ Catalyst UI Microsoft has been rolling out) severely restricts
  AppleScript automation; results may come back empty even when the
  classic UI would have found matches. The plugin still loads — the
  failure mode is "zero hits," not a crash.
- No StoreID dedup (Mac Outlook doesn't expose one); duplicate accounts
  with the same display name will produce duplicate hit rows.
- First run will trigger a macOS TCC prompt ("ComBridge wants to
  control Microsoft Outlook"). Approve in System Settings → Privacy &
  Security → Automation.

### Docs
- `LLM/plugins.md` Mac Outlook section updated to remove the "no
  search" caveat and document the new command.
- `LLM/troubleshooting.md` gained an entry covering "Mac outlook search
  returns zero hits" with the New-Outlook-for-Mac diagnosis and TCC
  permission check.
- FR `D:\Dev\FeatureRequests\ComBridge_FeatureRequests\FR_scripting_dx_and_outlook_search.md`
  followup section updated: the previously-deferred Mac equivalent is
  now shipped.

## [0.4.0] — script DX fixes + outlook search command

Addresses `D:\Dev\FeatureRequests\ComBridge_FeatureRequests\FR_scripting_dx_and_outlook_search.md`
in full. All four items shipped.

### Added
- **Wider script default references** — `ScriptHost.RunAsync` now adds
  `System.Text.RegularExpressions`, `System.Text.Json`, `System.Net.Http`,
  `System.Xml.ReaderWriter` + `System.Private.Xml`,
  `System.Diagnostics.Process`, and `System.Net.WebUtility` to the
  default reference set. User .csx files can now `using
  System.Text.RegularExpressions;` + `Regex.Replace(...)` without explicit
  `#r` directives. (FR item 1)
- **Documented default reference + import contract** — `LLM/scripting.md`
  gained a "Default reference set + import set" section enumerating every
  assembly the script can call into, every namespace auto-imported, how
  `#r` directives work, and a worked Office-Exception-ambiguity warning.
  (FR items 1, 2)
- **`outlook search` command** (Windows Outlook plugin) — recursive
  mail-content search across MAPI stores using DASL `Restrict` for
  speed. Flags: `--query`, `--store`, `--folder`, `--fields`, `--max`,
  `--since`, `--snippet`. Per-folder try/catch tolerates unscriptable
  stores; deduplicates stores by StoreID; snippet extraction via
  `Regex` (now in default refs). Live-tested against a multi-store
  Exchange/IMAP mailbox. (FR item 3)
- **Helpful "no plugins discovered" error message** — when subcommand
  dispatch finds no plugins, `Program.cs` now prints an explicit hint
  block covering the three real causes (binary not staged next to
  `plugins/`, plugin DLL naming, OS filter exclusion) instead of an
  empty `Available:` list. `LLM/troubleshooting.md` gained a new entry
  cross-referencing the subcommand-path symptom. (FR item 4)

### Changed
- `Program.cs` plugin-not-found error path now distinguishes
  "plugin name typo" vs "no plugins at all" — different messages, both
  actionable.

### Notes
- The Mac Outlook plugin does NOT yet have a `search` command. DASL
  `Restrict` is Windows-only; a Mac AppleScript equivalent using
  `whose`-clause filters would be ~100× slower and is deferred until
  there's demand.
- Roslyn `#r "Name"` directives for framework-resolvable assemblies
  already worked (Roslyn's default `ScriptMetadataResolver` ships with
  the host; `WithReferences` doesn't clear it). Now documented.

## [0.3.1] — full Mac Office coverage + CI + comprehensive docs sweep

### Added
- **`ComBridge.Mac.Common`** library — shared `Osascript` helper used by
  all Mac plugins. Extracted from Excel.Mac so Word.Mac / PowerPoint.Mac /
  Outlook.Mac aren't triplicating the same subprocess plumbing.
- **`ComBridge.Plugins.Word.Mac`** — AppleScript-backed Word for Mac
  plugin. Commands: `info`, `extract-text`, `doc-stats`. Same CLI name
  (`word`) as the Windows Word plugin.
- **`ComBridge.Plugins.PowerPoint.Mac`** — AppleScript-backed PowerPoint
  for Mac plugin. Commands: `info`, `list-slides`. Same CLI name
  (`powerpoint`).
- **`ComBridge.Plugins.Outlook.Mac`** — AppleScript-backed Outlook for
  Mac plugin (with documented limitations vs the Windows MAPI plugin —
  no Stores collection, thinner dictionary, "New Outlook for Mac"
  restrictions noted). Commands: `info`, `list-accounts`.
- **GitHub Actions CI** (`.github/workflows/build.yml`) — two jobs:
  - macOS runner: builds Core, CLI, Mac.Common, all 4 Mac plugins,
    smoke-tests `combridge list-plugins`.
  - Windows runner: builds Core (both TFMs), CLI (both TFMs), Mac
    plugins (proves they compile cross-platform). Windows plugins
    needing installed Office/SOLIDWORKS are NOT built (no app
    available on hosted runners; documented in workflow comments).
- **LLM docs full cross-platform sweep**:
  - `LLM/plugins.md` § "macOS plugins" with per-plugin specifics, AppleScript app names, what differs vs Windows, implementation notes
  - `LLM/authoring.md` § "macOS plugin pattern" — prescriptive template + reference layout + drop-in skeleton + other AppleScript-friendly apps
  - `LLM/troubleshooting.md` § "macOS / AppleScript issues" — TCC permission prompts, `osascript` slowness in loops, "Application isn't running" cause, New Outlook for Mac caveats, plugin-doesn't-load diagnostics
  - `LLM/build.md` — Mac build commands + per-OS plugin availability table
  - `LLM/workflow.md` task router — "Add a plugin for macOS" entry
  - `LLM/symbols.md` — Mac plugin deployment paths + Mac.Common library + symbol index
  - `LLM/README.md` defaults table — all 4 Mac plugins listed

### Architecture
- Plugin tree now categorized by platform:
  - 5 Windows plugins (SW + Office)
  - 4 Mac plugins (Office only — SolidWorks doesn't exist on macOS)
  - 1 shared library (Mac.Common)
- A single combridge bundle ships all 9 plugin folders side-by-side; the
  OS-supported ones load per machine.

### Migration
- No source changes needed for existing plugins.
- Existing v0.3.0 release tag remains valid; v0.3.1 adds Mac coverage
  + CI without breaking anything.

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
