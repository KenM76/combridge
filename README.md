# combridge

Cross-platform desktop-app automation host with a plugin model. Windows
plugins use COM; macOS plugins use AppleScript. One CLI contract
(`combridge <app> <command>`) for both.
Drop a plugin DLL into `plugins/<name>/` and the host gives you:

- ROT attach (or new-instance create) for that app's COM server
- Multi-instance session picker — `list-sessions` + `--session N|pid:NNNN|<title>|last` (default: MRU — the window the user was last focused on)
- A `run-script <file.csx>` command that compiles a C# script with the
  plugin's interop assemblies referenced and its globals injected
- Plus any plugin-specific commands (e.g. SW `active-doc`, Excel `dump-sheet`)

One exe + many plugins, instead of one exe per app.

> 📁 LLM-optimized references for this project live under [`LLM/`](LLM/).
> Human readers can use the prose docs (this README and
> [`PLUGIN_GUIDE.md`](PLUGIN_GUIDE.md)).

## Build

```powershell
cd <repo-root>          # the directory containing combridge.sln
dotnet build -c Release
```

The host exe lands in `src/ComBridge.Cli/bin/Release/net10.0-windows/combridge.exe`.
Plugin builds copy themselves into `<repo-root>/plugins/<Name>/` next to
the published exe (set up via the `CopyToPluginsRoot` MSBuild target in each
plugin csproj). To deploy, copy `combridge.exe` plus the `plugins/` folder
together.

**Build prerequisites per plugin:** the Excel plugin needs nothing extra
(its interop comes from NuGet). The SolidWorks plugin needs the SolidWorks
interop DLLs — on a normal machine with SolidWorks installed these are
found automatically via the Windows registry. If SW is installed somewhere
non-standard, set the path one of these ways (first wins):

```powershell
dotnet build -c Release /p:SolidWorksApiRedist="D:\SW2025\api\redist"   # one-off
$env:SOLIDWORKS_API_REDIST = "D:\SW2025\api\redist"                     # shell/system
# or copy src/plugins/ComBridge.Plugins.SolidWorks/paths.props.example
#         to paths.props (same folder) and edit it — persistent, gitignored
```

If the DLLs still can't be found, the build fails with a clear
`error COMBRIDGE001` that lists every location it tried.

## Usage

```text
combridge list-plugins
combridge <plugin> list-commands
combridge <plugin> list-sessions <outputFile>
combridge <plugin> [--session <sel>] <command> [args...] <outputFile>
```

The last positional argument is always the output file (use `-` for stdout).

### Flags

| Flag | Meaning |
|---|---|
| `--no-create` | Don't launch a new app instance if none is running. |
| `--session <sel>` | Pick a specific running instance. **Without it, the host attaches to the most-recently-focused session** — the window the user was last working in (via desktop Z-order). |

### `--session` selector forms

| Form | Example | Picks |
|---|---|---|
| (omitted) | `combridge solidworks active-doc -` | Most-recently-focused session (Z-order MRU) |
| `last` / `mru` / `recent` | `--session last` | Same as omitting — explicit MRU keyword (defensive against future default changes) |
| Pure digits | `--session 2` | 1-based index in MRU order: `1` = most recent, `2` = next-most, etc. |
| `pid:NNNN` | `--session pid:23456` | Win32 process ID — most deterministic |
| Any other string | `--session "Bracket"` | Case-insensitive substring of the active document/workbook title |

### SolidWorks (attach to running session)

```powershell
combridge solidworks list-sessions out.txt
combridge solidworks active-doc out.txt
combridge solidworks --session 2 active-doc out.txt
combridge solidworks run-script examples\sw_active_doc.csx out.txt
```

Globals exposed in scripts: `swApp`, `swDoc`, `swPart`, `swAssy`, `swDrawing`,
`swDocType`. Namespaces auto-imported: `SolidWorks.Interop.sldworks`,
`SolidWorks.Interop.swconst`, `SolidWorks.Interop.swcommands`.

SolidWorks is **never** silently launched — `AllowCreateNew` is `false` for
this plugin. Have SW open before running the bridge.

### Excel (attach or launch)

```powershell
combridge excel info out.txt
combridge excel list-sessions out.txt
combridge excel --session "Budget.xlsx" dump-sheet Sheet1 out.tsv
combridge excel run-script examples\excel_dump_active_sheet.csx out.tsv
```

Globals: `xlApp`, `xlBook`, `xlSheet`. Excel is launched fresh if no instance
is running (Office often skips ROT registration until a workbook is open).
Pass `--no-create` to force attach-only.

> **Office 365 caveat:** Excel's shared-instance shim consolidates workbooks
> back into one process even when a new `EXCEL.EXE` is spawned. Multi-instance
> Excel is supported in the code paths but rare in live use on default Office
> 365 installs.

### Word (attach or launch)

```powershell
combridge word list-sessions out.txt
combridge word info out.txt
combridge word --session "Report.docx" info out.txt
combridge word run-script my_word_script.csx out.txt
```

Globals: `wdApp` (`Word._Application`), `wdDoc` (`Word.Document?`).

### PowerPoint (attach or launch)

```powershell
combridge powerpoint list-sessions out.txt
combridge powerpoint info out.txt
combridge powerpoint run-script my_ppt_script.csx out.txt
```

Globals: `pptApp` (`PowerPoint._Application`), `pptPres` (`PowerPoint.Presentation?`),
`pptSlide` (`PowerPoint.Slide?`).

### Outlook (single MAPI session)

```powershell
combridge outlook list-sessions out.txt
combridge outlook info out.txt
combridge outlook run-script my_outlook_script.csx out.txt
```

Globals: `olApp` (`Outlook._Application`), `olNs` (`Outlook.NameSpace` —
the `"MAPI"` namespace, where folders/items live), `olExplorer`
(`Outlook.Explorer?` — the active window).

Outlook is single-instance by design (one MAPI session per user). The bridge
uses `GetActiveObject` only — no per-document ROT walking.

## Using ComBridge.Core from a third-party tool

If you're building a custom CLI/service/GUI that needs ROT attach, session
picking, or the Roslyn script host — but doesn't fit the CLI or plugin
shapes (e.g. it orchestrates SolidWorks + SwDocumentMgr in one process, or
has its own UX) — reference `ComBridge.Core.dll` directly as a library.

See [`CONSUMING_CORE.md`](CONSUMING_CORE.md) (humans) or
[`LLM/consuming.md`](LLM/consuming.md) (LLMs). Short version:

```xml
<PropertyGroup>
  <ComBridgeRoot>$([System.Environment]::GetEnvironmentVariable('COMBRIDGE_ROOT'))</ComBridgeRoot>
  <ComBridgeRoot Condition="'$(ComBridgeRoot)' == ''">D:\Dev\combridge</ComBridgeRoot>
</PropertyGroup>
<ItemGroup>
  <Reference Include="ComBridge.Core">
    <HintPath>$(ComBridgeRoot)\src\ComBridge.Core\bin\Release\net10.0-windows\ComBridge.Core.dll</HintPath>
  </Reference>
</ItemGroup>
<Import Project="$(ComBridgeRoot)\src\plugins\Common.Paths.props" />
```

`LLM/api.md` documents the stable API surface and stability tiers per type.

## Adding your own commands to an existing plugin

You can add custom named commands to any shipped plugin without forking
or recompiling it. Drop a `.csx` file into the plugin's `commands/`
folder and it becomes an invokable command:

```
plugins/SolidWorks/commands/my-export.csx     ← create this
```

```powershell
combridge solidworks list-commands
# →   active-doc      (plugin)    active-doc ...
#     my-export       (script)    my-export   (.csx: my-export.csx)   ← auto-discovered

combridge solidworks my-export out.txt
# runs your .csx against the SW plugin's globals
```

The script body uses the plugin's globals (`swApp`, `xlApp`, etc.) just like
`run-script` does. Built-in commands and the plugin's own typed commands
take precedence on name collision — your scripted command can never
accidentally shadow them.

See [`LLM/extending.md`](LLM/extending.md) for the full convention,
including "why not DLL-based sub-plugins" and the criteria for adding
that later if needed.

## Adding a new plugin

See [`PLUGIN_GUIDE.md`](PLUGIN_GUIDE.md) (humans) or
[`LLM/plugins.md`](LLM/plugins.md) (LLMs). Short version:

1. New class library that references `ComBridge.Core`
2. Add the app's interop assemblies (NuGet PIA, or `<Reference>` with HintPath)
3. Implement `IComBridgePlugin` — name, ProgIDs, globals type, commands, optional `DescribeInstance`
4. Add a `CopyToPluginsRoot` target so the build deploys to `plugins/<Name>/`
5. Drop into the .sln and rebuild

## Layout

```
combridge/
  combridge.sln
  Directory.Build.props          # shared TFM = net10.0-windows, nullable, etc.
  .gitignore                     # excludes bin/, obj/, paths.props (live override)
  README.md                      # this file
  PLUGIN_GUIDE.md                # author guide for new plugins
  CONSUMING_CORE.md              # library-mode guide for third-party tools
  LLM/                           # LLM-optimized docs (dense, structured)
    README.md                    #   entry index + top-level facts
    api.md                       #   ComBridge.Core public surface + stability tiers
    cli.md                       #   CLI grammar + exit codes
    plugins.md                   #   per-plugin specifics + new-plugin recipe
    build.md                     #   pitfalls table + csproj boilerplate
    paths.md                     #   machine-specific path resolution chain
    consuming.md                 #   library mode for third-party tools
    symbols.md                   #   symbol → file index
  src/
    ComBridge.Core/              # IComBridgePlugin, ROT, Roslyn host, loader, session picker
    ComBridge.Cli/               # combridge.exe entry + dispatch
    plugins/
      Common.Paths.props         #   shared path-resolution + COMBRIDGE001 validation
      ComBridge.Plugins.SolidWorks/
        paths.props.example      #     committed template; live paths.props is gitignored
      ComBridge.Plugins.Excel/
      ComBridge.Plugins.Word/
      ComBridge.Plugins.PowerPoint/
      ComBridge.Plugins.Outlook/
  plugins/                       # build output — what the exe loads at runtime
    SolidWorks/
    Excel/
    Word/
    PowerPoint/
    Outlook/
  examples/                      # ready-to-run .csx scripts
```
