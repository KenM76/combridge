# Extending Plugins — Per-User / Per-Site Commands

`combridge` ships with one global extension point: drop a `.csx` file into
`plugins/<PluginName>/commands/` and it becomes a named command on that
plugin. No DLL, no recompile, no fork of the plugin.

This is the **per-user / per-site** extension story. It's deliberately
lightweight — see "Why not DLL-based sub-plugins?" at the bottom for the
shape that's been considered and (deliberately) not implemented.

## The convention

```
plugins/SolidWorks/
  ComBridge.Plugins.SolidWorks.dll       ← shipped plugin (don't touch)
  SolidWorks.Interop.sldworks.dll
  ...
  commands/                              ← create this folder
    my-export-cam.csx                    ← becomes `combridge solidworks my-export-cam`
    site-bom-report.csx                  ← becomes `combridge solidworks site-bom-report`
    weekly-revision-stamp.csx
```

`combridge solidworks list-commands` enumerates everything in one table:

```
  run-script             (built-in)  run-script <scriptFile.csx>
  list-sessions          (built-in)  list running instances of this plugin's app
  active-doc             (plugin)    active-doc   (prints title + type of active document)
  my-export-cam          (script)    my-export-cam   (.csx: my-export-cam.csx)
  site-bom-report        (script)    site-bom-report   (.csx: site-bom-report.csx)
  weekly-revision-stamp  (script)    weekly-revision-stamp   (.csx: weekly-revision-stamp.csx)
```

The `(plugin)` / `(script)` annotation tells you where a command came from.

## How it works internally

`PluginLoader.GetScriptedCommands(plugin)` scans
`<plugin-deploy-dir>/commands/*.csx` and wraps each file in a
`ScriptedCommand : IBridgeCommand`. The CLI dispatcher in `Program.cs`
checks command names in this order:

1. **Built-ins** (`run-script`, `list-sessions`) — always win
2. **Plugin's typed `Commands`** — the plugin author's intentional API
3. **Scripted commands from `commands/*.csx`** — per-user extensions

A scripted command CANNOT shadow a built-in or a typed command. This is
intentional: the host and the plugin author remain authoritative; user
extensions can only add, not override.

Inside a scripted command, the script body runs through the same Roslyn
host as `run-script`. The plugin's globals are available
(`swApp`/`swDoc`/`xlApp`/`olApp`/etc.), `Console.Out` is redirected to
the bridge's output file, the same encoding/ALC/dynamic refs fixes
apply. Everything in `LLM/scripting.md` works inside `commands/*.csx`
verbatim.

## Conventions

| Topic | Rule |
|---|---|
| **Filename → command name** | Strip the `.csx` extension. `my-export.csx` → `my-export`. Use lowercase + hyphens by convention. |
| **Subdirectories** | NOT scanned recursively. Only top-level `*.csx` under `commands/` becomes commands. Put helpers/shared code in a different folder if you need it. |
| **Case sensitivity** | Command lookup is case-insensitive (`combridge solidworks DOC-SUMMARY` works for `doc-summary.csx`). |
| **Name collisions among scripted commands** | The first match in directory enumeration wins. Don't ship two `commands/foo.csx`. |
| **Name collisions with built-ins or typed commands** | Built-in/typed wins. The scripted version is silently ignored. |
| **Empty `commands/` folder** | Treated as "no scripted commands" — no error. |
| **Missing `commands/` folder** | Same — silently no scripted commands. |

## Privacy / gitignore

The deployed `plugins/<Name>/commands/` lives under the gitignored
top-level `/plugins/` folder. So per-user scripts can't accidentally end
up in the combridge repo via `git add .`. For sharing site-specific
commands within an organization:

- Keep them in a separate company repo / share
- Distribute via your deployment tooling (copy into the user's
  `plugins/<App>/commands/` as part of installing combridge for them)
- Or just hand them around as `.csx` files — they're plain text

## Examples — promoting an existing `.csx` to a scripted command

Any file under `examples/` is a candidate. Suppose you use
`examples/sw_iterate_components.csx` regularly:

```powershell
# Before (verbose):
combridge solidworks run-script examples\sw_iterate_components.csx out.txt

# Promote to a scripted command:
copy examples\sw_iterate_components.csx plugins\SolidWorks\commands\iter-comps.csx

# After (terse):
combridge solidworks iter-comps out.txt
```

The script's existing globals usage (`swAssy`, `swDoc`, etc.) works
identically.

## Versioning concerns

A scripted command depends on:
- The plugin's globals shape (`SwGlobals`, `ExcelGlobals`, etc.)
- The plugin's `ScriptUsings` (auto-imported namespaces)
- The plugin's `ScriptReferences` (compile-time available interop)

These can in principle change between combridge versions. The stability
tiers in `LLM/api.md` cover them: `SwGlobals`/`ExcelGlobals`/etc. are
"stable shape, evolving members" — names + types of existing fields
won't change; new fields may be added. So a scripted command that uses
`swApp.ActiveDoc` today keeps working when a future combridge adds
`swApp.SomeNewProperty`.

If you want defensive guarantees, pin combridge to a specific commit
hash or release tag in your deployment.

---

# Why not DLL-based sub-plugins? (Not on the roadmap)

The natural "more powerful" alternative is to let users ship typed,
compiled sub-plugins that reference the parent plugin's assembly and
add their own `IBridgeCommand` implementations. This was considered and
deliberately not built.

## What DLL sub-plugins would add over `.csx`

| Capability | `.csx` (Shape A) | Hypothetical DLL sub-plugin (Shape B) |
|---|---|---|
| Strong typing (IntelliSense in an IDE) | partial (Roslyn-script-level) | full |
| NuGet dependencies | only what the parent plugin exposes | the sub-plugin can declare its own |
| Multi-file structure | no — one script per command | yes — full project layout |
| Cross-invocation state | no (combridge.exe exits between calls; cold-start every time) | same — combridge.exe lifecycle doesn't change |
| Compile-time error checking | runtime (in `run-script` wrapper) | full compile-time |
| Distribution as a binary (without source) | source is the .csx file (obfuscation only) | distribute compiled .dll only |
| Per-command versioning | shared with the parent plugin | sub-plugin can have its own SemVer |

## What it would cost to build

Roughly 300 lines plus a non-trivial set of doc updates. The key
architectural pieces:

- **New contract**: `IComBridgeSubPlugin` with `ParentPluginName` and
  `Commands` properties.
- **AssemblyLoadContext bridging**: the sub-plugin's ALC needs to defer
  to the parent plugin's ALC for the parent's types (otherwise the
  globals object created by the parent plugin can't be cast to the
  globals type as the sub-plugin sees it — the same ALC identity
  mismatch that bit us in Roslyn scripting and that we solved with
  `InteractiveAssemblyLoader.RegisterDependency`).
- **Discovery convention**: `plugins/<Parent>/extensions/<Sub>/<Sub>.dll`
  vs. an explicit manifest. Both have trade-offs.
- **Version compatibility**: the sub-plugin compiles against a specific
  combridge + parent plugin version. Need either a stability promise
  beyond what `LLM/api.md` documents, or a runtime version check, or
  both.
- **List-commands integration**: a `(subplugin: Acme.SwCam)` annotation
  to show provenance.

It's the kind of feature that's straightforward to implement once but
expensive to maintain across versions (every parent-plugin API change
risks breaking pre-compiled sub-plugins in ways that .csx wouldn't
break, because .csx is recompiled every run).

## What would warrant building it

**Real-world triggers** — if any of these come up, the cost-benefit
flips and Shape B becomes worth building. Until then, Shape A covers
the use case better:

1. **A user wants to distribute a proprietary sub-plugin as a binary
   only** (no source visible). `.csx` is plain text; even obfuscation
   doesn't really hide intent. Real protection requires compiled DLLs.
2. **A sub-plugin needs NuGet packages the parent plugin doesn't
   include**, AND those packages don't work loaded ad-hoc from a script
   (e.g., source generators, T4 templates, anything compile-time-y).
3. **A sub-plugin grows beyond ~500 lines** with multiple related
   commands sharing typed helper classes. At that size, keeping it as
   parallel `.csx` files with copy-pasted helpers becomes painful.
4. **A team wants full IDE IntelliSense + refactoring on their custom
   commands**, not just Roslyn-script-level support.
5. **A sub-plugin needs to register with combridge for non-command
   purposes** — e.g., add a custom selector form, register a moniker
   pattern at runtime, hook into the COM-attach pipeline. (None of this
   is possible with either Shape A OR Shape B as proposed; it'd require
   architectural changes beyond either.)

## What would NOT warrant building it

- "I want to share commands across my team" — `.csx` files are easier
  to share than DLLs (just copy the file).
- "I want my commands in source control" — `.csx` files live in source
  control fine; per-user `plugins/<App>/commands/` is gitignored only
  because the parent `/plugins/` folder is gitignored at the repo level
  (it's the build-output deploy target). You can keep a separate repo
  of per-site `.csx` commands and copy them into place during install.
- "I want compile-time checks" — Roslyn's diagnostics (CS0xxx codes)
  fire when `run-script` invokes a malformed .csx, and the error is
  printed before any side effects run. That's almost-but-not-quite
  compile-time enough for most cases.

## Decision rule

When the first real use case from the "would warrant" list above shows
up — write Shape B then. Until then, Shape A is the right level of
complexity for the demonstrated need.
