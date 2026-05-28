# combridge CLI Grammar

## EBNF

```
invocation        = "combridge" , flag* , (
                      "list-plugins"
                    | plugin-name , flag* , subcommand , flag*
                    ) ;

subcommand        = "list-commands"
                  | "list-sessions" , output-arg
                  | command-name , [ command-arg ]* , output-arg ;

flag              = "--no-create"
                  | "--session" , whitespace , selector
                  | "--session=" , selector ;

selector          = digits                        (* 1-based index in MRU order *)
                  | "pid:" , digits               (* Win32 PID                  *)
                  | "last" | "mru" | "recent"     (* explicit MRU keyword (CI)  *)
                  | other-string ;                (* title substring (CI)       *)

output-arg        = path | "-" ;                  (* "-" routes to stdout *)
```

Argument parsing in `Program.cs::Main` is order-tolerant: flags can appear anywhere, including between the plugin name and the command. Flags are stripped from `args` before dispatch.

## Examples (canonical)

```
combridge list-plugins
combridge solidworks list-commands
combridge solidworks list-sessions out.txt
combridge solidworks active-doc out.txt
combridge solidworks --session 2 active-doc out.txt          # 2nd most-recently-focused
combridge solidworks --session last active-doc out.txt       # explicit "most recent" (same as omitting --session)
combridge solidworks --session pid:23456 run-script foo.csx out.txt
combridge solidworks --session "Bracket" run-script foo.csx out.txt
combridge excel info -                                       # default → MRU Excel; stdout
combridge --no-create excel info out.txt                     # forbid launch
```

## Exit codes

| Code | Meaning | Source |
|---|---|---|
| 0 | success | normal |
| 1 | unhandled `Exception` in `Main` | catch-all in `Program.Main` |
| 2 | script file not found | `ScriptHost.RunAsync` |
| 3 | Roslyn compilation error(s) | `ScriptHost.RunAsync` |
| 4 | script ran but threw (`state.Exception`) | `ScriptHost.RunAsync` |
| 5 | host exception during script run (post-compile) | `ScriptHost.RunAsync` |
| 6 | could not connect / no session matched selector | `Program.Main` (RotHelper or SessionPicker.Resolve) |
| 64 | usage error (unknown plugin/command, missing output file) | `Program.Main` |

## Output-file convention

The LAST positional argument (after flag stripping) is the output file. Use `-` for stdout. The host opens `new StreamWriter(outputFile)` (or `Console.Out`) and passes it to `IBridgeCommand.RunAsync`. Inside `run-script`, the script's `Console.WriteLine` is also redirected to this writer.

## Default session ordering — MRU (most-recently-used)

When `--session` is omitted, combridge picks the **most recently focused**
session — the one whose top-level window is highest in Windows' global
Z-order (excluding the terminal/IDE the user typed `combridge` into,
since its PID doesn't match any plugin's session). After clicking SW
window B then switching to a terminal to run `combridge solidworks ...`,
the bridge attaches to **B**.

Implementation lives in `SessionPicker.Enumerate` — see `RankByZOrder`.
`SessionInfo.Index` is assigned 1-based per the MRU-sorted order, so
`--session 1` is MRU, `--session 2` next-most-recent, etc. Sessions
whose process has no visible top-level window (minimized to tray,
headless) sort to the end and preserve ROT discovery order among
themselves.

### Selector forms

| Form | Picks |
|---|---|
| (no `--session`) | Most-recently-focused session (Z-order #1 among matching PIDs) |
| `last` / `mru` / `recent` (CI) | Same as omitting `--session` — explicit MRU keyword, defensive against future default changes |
| pure digits, e.g. `2` | 1-based index in MRU order (`1` = most recent, `2` = next, etc.) |
| `pid:NNNN` | Match by Win32 PID |
| anything else | Case-insensitive substring of title (or full description) |

### Behavior matrix: --session × ROT state × AllowCreateNew

| --session given | Sessions visible | AllowCreateNew | Result |
|---|---|---|---|
| no | ≥ 1 | any | **MRU session** (top of Z-order among matching PIDs) |
| no | 0 | true | Create new instance via `Type.GetTypeFromProgID` |
| no | 0 | false (or `--no-create`) | Exit 6 |
| yes | ≥ 1 + selector matches | any | Picked instance |
| yes | ≥ 1 + selector misses | any | Exit 6, list available |
| yes | 0 | any | Exit 6 (does NOT create even if AllowCreateNew) |

Rationale: passing `--session` is an explicit "I want one of the running ones"; falling through to launch a new instance would be surprising.

## Built-in commands handled before COM attach

`list-plugins`, `list-commands`, `list-sessions` do NOT trigger ROT attach beyond what `list-sessions` does internally (read-only enumeration). Other commands (including unknown ones that get rejected) attach first only after dispatch validation.

## list-sessions output format

```
Running <plugin-name> sessions:
  #1  pid=NNNN  title=<title-or-(none)>
  #2  pid=NNNN  title=<title-or-(none)>

Select with --session N (1=MRU) | --session pid:NNNN | --session <title> | --session last
Default with no --session = MRU (most-recently-focused window).
```

## `list-commands` output format

```
  run-script      (built-in)  run-script <scriptFile.csx>
  list-sessions   (built-in)  list running instances of this plugin's app
  active-doc      (plugin)    active-doc   (prints title + type of active document)
  my-export       (script)    my-export   (.csx: my-export.csx)
```

Category annotations:
- **`(built-in)`** — provided by combridge itself, available on every plugin
- **`(plugin)`** — defined by the plugin's `IComBridgePlugin.Commands` (typed C# `IBridgeCommand`)
- **`(script)`** — auto-discovered `.csx` file in `<plugin-deploy-dir>/commands/`. See `LLM/extending.md` for the convention.

Command lookup precedence on name collision: built-in beats plugin beats script. A scripted command can never shadow a built-in or a typed command.

If no sessions: `(no running <plugin-name> sessions in the ROT)`.

Order is MRU (most-recently-focused first). `#1` is always the session whose
window the user last interacted with; sessions whose process has no visible
top-level window fall to the end. Entries where `DescribeInstance` returns
both null PID AND empty title are filtered out by the dead-binding rule
in `SessionPicker.Enumerate` (they aren't shown).
