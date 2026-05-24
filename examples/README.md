# `examples/` — ready-to-run `.csx` scripts

Each `.csx` runs against a single plugin's globals via:

```powershell
combridge <plugin> run-script examples\<file>.csx <outputFile>
# (use - for stdout)
```

Scripts are independent — copy any one, edit, and run. For broader
recipe coverage (cookbook-style snippets you can compose), see
`LLM/scripting.md`.

## Index

### SolidWorks (`combridge solidworks ...`)
| File | What it does |
|---|---|
| `sw_active_doc.csx` | Print active doc title, path, type (smoke test) |
| `sw_iterate_components.csx` | List top-level components of an open assembly with paths, configs, suppression |
| `sw_feature_walk.csx` | Walk every feature on the active doc — read-only, safe |

### Excel (`combridge excel ...`)
| File | What it does |
|---|---|
| `excel_dump_active_sheet.csx` | Dump the active sheet's used range as TSV |
| `excel_iterate_workbooks.csx` | Enumerate every open workbook + its worksheets |
| `excel_find_value.csx` | Search the active sheet's used range for cells matching a literal |
| `excel_write_new_workbook.csx` | Create a new workbook, write a 10×3 table, save as `.xlsx` |

### Word (`combridge word ...`)
| File | What it does |
|---|---|
| `word_doc_stats.csx` | Word / character / line / page / paragraph counts |
| `word_extract_text.csx` | Dump the active document's plain text |
| `word_find_replace_template.csx` | Doc-wide find-and-replace (edit the two strings first) |

### PowerPoint (`combridge powerpoint ...`)
| File | What it does |
|---|---|
| `powerpoint_list_slides.csx` | List every slide with title + shape count |
| `powerpoint_export_pdf.csx` | Export active presentation to PDF |

### Outlook (`combridge outlook ...`)
| File | What it does |
|---|---|
| `outlook_unread_inbox.csx` | List unread Inbox items (sender, time, subject) |
| `outlook_today_calendar.csx` | List today's calendar appointments |
| `outlook_list_stores.csx` | List every configured store/account + Inbox item counts |

## Conventions

- **Edit-then-run idioms** — examples that need a target path or search string
  declare a `const string` at the top with an `← edit me` comment. Adjust
  in-file before running.
- **Null guards** — every example that uses a per-document global
  (`xlBook`, `wdDoc`, `pptPres`, `swAssy`, etc.) checks for null and
  exits with code 2 if the app has nothing open.
- **Roslyn-script-specific gotchas**:
  - `using var x = ...` declaration form is NOT supported — use
    `using (var x = ...) { ... }` block form.
  - Top-level `var` declarations hoist but assignments don't —
    keep order natural.
  - `dynamic` works (Microsoft.CSharp + DynamicAttribute + CallSite
    refs are in the default script options).
- **Don't manually `Marshal.ReleaseComObject` globals** — the bridge owns
  the top-level RCW; aggressive release breaks subsequent calls.
- **SW only**: do NOT call mutation APIs (`ForceRebuild3`,
  `EditFeature`, `SetSuppression2` in a loop, `UpdateCutList` on nested
  folders) inside an iteration. Several SW APIs have non-obvious crash
  modes — see `C:\personal_rag\solidworks\` lessons before writing
  anything beyond read-only iteration.

## Writing your own

Use these as templates. The plugin's globals are documented in
`LLM/plugins.md` § "<plugin> plugin" — that's the authoritative list of
what you can reach from a script. For pattern catalogs that go beyond
these examples, see `LLM/scripting.md`.
