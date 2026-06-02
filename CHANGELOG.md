# Changelog

All notable changes to this project will be documented in this file.
Format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/);
versions follow [SemVer](https://semver.org/spec/v2.0.0.html).

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
