# LLM Workflow — Where to Start

If you're an LLM working on this project for the first time in this
session, **read this file first**, then jump to the doc that matches the
user's task. Don't grep blindly across the whole repo — the LLM docs
are designed so the right file gives you everything you need.

## Bootstrap: read once per session

These three files give you the gestalt. Read them in this order:

1. **`LLM/README.md`** (~150 lines) — index, file map, top-level facts,
   defaults reference table for all shipped plugins.
2. **`LLM/api.md`** (~200 lines, top sections) — public API surface +
   stability tier table. You need to know what's stable before changing
   anything in `ComBridge.Core`.
3. **This file** (`LLM/workflow.md`).

That's the minimum bootstrap. ~10 minutes for a frontier model.

## Task router

Match the user's request to one of these patterns, then read the
indicated docs (in order) before responding:

### "Run X in [app] from my CLI" — use existing plugin

User wants to do something with Excel/Word/etc. from a one-liner.

1. `LLM/cli.md` for the command grammar
2. `LLM/plugins.md` § "<app> plugin" for the plugin's commands + selectors
3. Respond with the exact `combridge <plugin> <command> ...` invocation

### "Add my own named command to [app]'s plugin (without forking it)"

User wants a custom command on an existing plugin — e.g.
`combridge solidworks my-export-cam` — sourced from their own `.csx`.

1. **`LLM/extending.md`** — explains the `plugins/<Name>/commands/*.csx`
   auto-discovery convention. The whole answer is one file drop.
2. `LLM/scripting.md` for the script body itself (per-plugin idioms).
3. Built. No code change, no plugin fork. Tell the user to verify with
   `combridge <plugin> list-commands` — the new command should appear
   labeled `(script)`.

### "Write me a [.csx] script that does X in [app]"

User wants a Roslyn script that uses an existing plugin's globals.

1. `LLM/scripting.md` for per-plugin recipes — find the closest match
2. `LLM/plugins.md` § "<app> plugin" for the globals' exact types + commands
3. If the recipe doesn't exist, compose from the recipes' building blocks
4. Test mentally for the Roslyn-script-specific traps:
   - `using var x = ...` declaration form: **NOT supported** — use `using (var x = ...) { ... }` block form
   - Top-level `var` declarations hoist but assignments don't — order matters
   - `dynamic` is supported (default refs include Microsoft.CSharp)
   - Top-level statements only — no `class Program { static void Main }` wrapper

### "Add a plugin for [app]"

User wants combridge to support a new COM-automation app.

1. **`LLM/authoring.md`** — read it ALL. This is the prescriptive guide.
2. Figure out which of the four discovery patterns the app uses (A/B/C/D in authoring.md § Step 1)
3. If `app` is in authoring.md's "Worked examples" section, use that as the skeleton
4. Cross-reference `LLM/build.md` for csproj pitfalls
5. Cross-reference `LLM/paths.md` if the interop is at a machine-specific path
6. Follow the verification checklist in authoring.md before declaring done

For apps NOT in authoring.md's worked examples, the four-pattern table
in Step 1 is the decision aid. Default to **Pattern A** (file monikers
+ Application ascent) unless there's evidence otherwise.

### "I got an error: [error message]"

User pasted an error.

1. **`LLM/troubleshooting.md`** — find the error code or symptom in the catalog
2. If not in troubleshooting.md, check `LLM/build.md` § Pitfalls (csproj-specific)
3. Follow the troubleshooting entry's "Fix" exactly — don't improvise
4. If it cross-references a `C:\personal_rag\...` lesson, read that for context

### "How do I consume Core as a library?"

User is building a third-party tool that needs ComBridge.Core directly,
not via the CLI.

1. **`LLM/consuming.md`** — full guide
2. `LLM/api.md` § "Stability contract" — verify what they want to depend on is stable
3. `LLM/paths.md` § "External consumers" if they also need the path-resolution chain

### "Why does X behave this way?" — design/architecture questions

User wants to understand the design.

1. `LLM/README.md` § "Top-level facts" — usually has the answer
2. `LLM/api.md` for behavior of specific types
3. `LLM/cli.md` § "Behavior matrix" for --session × ROT × AllowCreateNew
4. If still unclear, search code via `LLM/symbols.md` → file location

### "Find me [symbol/class/method]"

User wants to locate code.

1. **`LLM/symbols.md`** — symbol-to-file index, no grep needed

### "SolidWorks did X weird"

SW-specific. Many empirical findings live in `C:\personal_rag\solidworks\`
(host-specific lesson location).

1. `LLM/plugins.md` § "SolidWorks plugin" for SW plugin specifics
2. **`C:\personal_rag\solidworks\index.md`** — empirical lessons (crashes, quirks). Grep this folder for any SW API symbols in the user's question.
3. The canonical SW API reference if available on the host: `C:\sw_api_docs\rag_optimized\sldworks_methods_v3_llm.rag` (method signatures), `swconst_enums.txt` (enum values)

### "Office is doing X weird"

Office quirks live in `C:\personal_rag\claude_code\` (since they're
relevant to anything using Office COM, not just combridge).

1. `LLM/plugins.md` § "Excel plugin" / "Word plugin" / etc.
2. `LLM/troubleshooting.md` § "Office shared-instance behavior" for the
   most common surprise
3. `C:\personal_rag\claude_code\lesson_20260521_*.md` — three lessons
   covering ROT walking vs GetActiveObject, dispinterface vs co-class,
   Office 365 shared-instance shim

## File map (quick reference)

| File | Purpose | When to read |
|---|---|---|
| `LLM/README.md` | Entry index, defaults table, top-level facts | Bootstrap (always read first) |
| `LLM/workflow.md` | This file — task router | Bootstrap (always read second) |
| `LLM/api.md` | Public API surface + stability tiers | Bootstrap (always skim); deep-read when changing Core |
| `LLM/cli.md` | CLI grammar, exit codes, selector behavior | When writing CLI invocations |
| `LLM/plugins.md` | Per-plugin specifics for shipped plugins | When using a shipped plugin |
| `LLM/scripting.md` | Per-plugin .csx recipe cookbook | When writing scripts |
| `LLM/authoring.md` | How to build a NEW plugin (prescriptive) | When adding plugin support for a new app |
| `LLM/build.md` | Build pitfalls + csproj boilerplate | When a build fails or a new csproj is needed |
| `LLM/paths.md` | Machine-specific path resolution chain | When an interop DLL isn't found |
| `LLM/troubleshooting.md` | Error → cause → fix catalog | When an error fires (any phase) |
| `LLM/consuming.md` | Library-mode (Core as a third-party dep) | When the user is building their own tool |
| `LLM/symbols.md` | Symbol → file index | When locating code |

## External knowledge (machine-specific, if available)

These files are on the host machine and may not be available
everywhere — they're highly valuable when they ARE present.

| Path | Contents |
|---|---|
| `C:\personal_rag\claude_code\` | Lessons about Claude Code itself + COM/Roslyn/Office quirks discovered while building this and related projects |
| `C:\personal_rag\solidworks\` | Empirical lessons about the SW COM API — crashes, undocumented behavior, quirks |
| `C:\personal_rag\dxf\` | DXF post-processing knowledge (mostly for SW DXF export workflows) |
| `C:\sw_api_docs\rag_optimized\` | Canonical SOLIDWORKS API reference (method signatures, enum values) |
| `C:\tax_rag\` | Canadian tax law RAG — not combridge-related but on the same machine |

If you can read these, do so when relevant. The `personal_rag` lessons
are NON-OBVIOUS findings that won't be in any canonical docs — they
save real time.

## When you're done with a task

Before declaring done, verify against this list:

1. **The user's question is fully answered** — not just "here's how
   you'd approach it" but "here's the exact command/script/code."
2. **If you built or modified a plugin**: walked the verification
   checklist in `LLM/authoring.md` § "Verification checklist"
3. **If you found a non-obvious finding** that took >15 minutes to
   derive — write a lesson to `C:\personal_rag\claude_code\` (or
   `\solidworks\` if SW-specific). Update the indexes.
4. **If you changed a public API in Core** — update `LLM/api.md`
   stability tier and bump if breaking.
5. **If you added a new doc concept** — update `LLM/symbols.md`'s
   Documentation table and `LLM/README.md`'s file map.

## What to avoid

- **Don't grep across the entire repo when an LLM/*.md file would
  answer the question.** Grep is for things the LLM docs missed.
- **Don't add per-plugin --flags to the CLI.** The selector grammar
  in `LLM/cli.md` is the cross-plugin contract.
- **Don't break the stability tier promises in `LLM/api.md`** without
  bumping the tier and noting the break.
- **Don't write code with hardcoded machine-specific paths.** Use
  `Common.Paths.props` chain (`LLM/paths.md`).
- **Don't commit `paths.props`** (the live override file). Only
  `paths.props.example` is tracked.
- **Don't use `dynamic` in ScriptHost or Core code** — keep dynamic
  dispatch to user scripts. Core/ScriptHost should be statically
  typed.
- **Don't write lessons inside this repo.** Personal-machine lessons
  live in `C:\personal_rag\` so they survive across projects. The
  repo's docs are for the project's stable architecture; lessons are
  for empirical surprises.
