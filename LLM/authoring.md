# Building a New Plugin

This is the **prescriptive** guide for writing a combridge plugin against
an app that isn't shipped in this repo (AutoCAD, Inventor, Acrobat,
Photoshop, BricsCAD, Visio, Project, etc.).

If you're an LLM tasked with adding a plugin and the app isn't already in
`src/plugins/`, **start here, follow the steps in order, and check the
verification list at the end before declaring victory.**

If the app you want to automate is one we already ship (SolidWorks, Excel,
Word, PowerPoint, Outlook), read `LLM/plugins.md` for that plugin's
specifics — don't re-derive the work.

## Step 0 — Sanity check: is automation even possible?

Before writing code, verify:

1. **The app has a COM automation server.** Quickest test, run in PowerShell 5.1:
   ```powershell
   $obj = New-Object -ComObject "<App>.Application"  # e.g. "AutoCAD.Application"
   ```
   If this throws "Cannot find an object with CLSID …" or similar, the app
   doesn't expose COM. **Stop.** Lite/Reader/student versions often don't.
   Examples:
   - **AutoCAD LT** (any version): no COM. ProgID isn't registered.
   - **Adobe Reader** (free): some versions; Acrobat Pro has more.
   - **Visio Viewer**: no.
2. **It has an interop assembly** somewhere — NuGet PIA, GAC, vendor SDK
   folder. Without typed bindings, you'd have to use raw `dynamic`/IDispatch
   for everything which loses script ergonomics.
3. **It exposes the data you need** through that COM surface. Some apps
   have automation but a thin surface (e.g. Notepad++ has a scripting
   surface but not full COM).

If any of those fail, the plugin can't be built — combridge bridges to
COM, not to UI automation, not to in-process scripting, not to web APIs.

## Step 1 — Identify the app's automation pattern

Every COM app falls into one of these four discovery patterns. Knowing
which one your target follows determines `RotMonikerPatterns` and
`TryExtractRoot` content.

| Pattern | Examples | ROT registration | RotMonikerPatterns | TryExtractRoot |
|---|---|---|---|---|
| **A. Document-based, file monikers** | Excel, Word, PowerPoint, AutoCAD (full), Inventor, Visio, Photoshop (sort of) | One file moniker per open document (path string in ROT). NO class moniker. | Regex on file extensions: `\.(xlsx|...)$` | Bind moniker → Document RCW → ascend `.Application` via dynamic |
| **B. Multi-instance, custom-format monikers** | SolidWorks | Custom per-process moniker like `SolidWorks_PID_<pid>`. NO class moniker. | Match the custom format: `^SolidWorks_PID_\d+$` | Identity (the bound object IS the application) |
| **C. Single-instance class-moniker** | Outlook, some Adobe apps | Standard class moniker reachable via `oleaut32!GetActiveObject(ProgID)`. May not appear in ROT enumeration with a discoverable name. | Empty `Array.Empty<string>()` — rely on `GetActiveObject` fallback alone | Identity (not called — no ROT matches) |
| **D. Version-suffixed multi-ProgID** | AutoCAD (`AutoCAD.Application.24`), Visio (`Visio.Application.16`), some Inventor builds | Same as A or B underneath; just need multiple ProgID entries in case the user has multiple versions installed | Same as A or B | Same as A or B |

**How to tell which pattern your app uses:**

Walk the ROT directly with the diagnostic probe pattern. Quickest path is
to write a tiny .csx for an EXISTING plugin (any one) and `Console.WriteLine`
the contents of `IRunningObjectTable.EnumRunning`. Or use PowerShell to
poke `GetActiveObject` and check the ROT display names. See
`LLM/troubleshooting.md` § "Don't know which discovery pattern an app uses?"
for the diagnostic code.

## Step 2 — Locate the interop assembly

Three common sources, in preference order:

1. **NuGet PIA** (`Microsoft.Office.Interop.<App>`) — preferred for Office.
   Pros: version-pinned in csproj; portable across machines.
   Cons: needs `EmbedInteropTypes=false` + `CopyLocalLockFileAssemblies=true`
   (see `LLM/build.md` pitfalls).

2. **GAC HintPath** (`C:\Windows\assembly\GAC_MSIL\<App.Interop>\<ver>__<token>\`)
   — preferred for Office apps where NuGet doesn't cover all transitive deps
   (this is what Word/PowerPoint/Outlook do; `office.dll` always needs this).
   Pros: matches what's installed; same path on every Office machine.
   Cons: hardcoded GAC path, won't work on machines without Office.

3. **Vendor install dir HintPath** (e.g. `C:\Program Files\Autodesk\AutoCAD <ver>\`)
   — needed for non-Office apps that ship interop with their install.
   Pros: only path that works for these apps.
   Cons: machine-specific path → use `Common.Paths.props` resolution chain
   (registry probe + env var + default + COMBRIDGE001 validation). See
   `LLM/paths.md` for the full pattern.

Find the interop on the target machine first. For an unknown app:

```powershell
# Search common locations
Get-ChildItem -Path "C:\Windows\assembly\GAC_MSIL" -Filter "*<App>*" -Recurse -ErrorAction SilentlyContinue
Get-ChildItem -Path "C:\Program Files\<Vendor>" -Filter "*.Interop.dll" -Recurse -ErrorAction SilentlyContinue
Get-ChildItem -Path "C:\Program Files\<Vendor>" -Filter "*.tlb" -Recurse -ErrorAction SilentlyContinue
```

If only a `.tlb` (type library) exists with no managed wrapper, you'd need
to generate one with `tlbimp.exe` first — out of scope for this guide but
not blocked.

## Step 3 — Create the directory + csproj

Directory: `src/plugins/ComBridge.Plugins.<AppName>/`

`ComBridge.Plugins.<AppName>.csproj` skeleton:

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Library</OutputType>
    <RootNamespace>ComBridge.Plugins.<AppName></RootNamespace>
    <AssemblyName>ComBridge.Plugins.<AppName></AssemblyName>
    <PlatformTarget>x64</PlatformTarget>
  </PropertyGroup>

  <ItemGroup>
    <ProjectReference Include="..\..\ComBridge.Core\ComBridge.Core.csproj">
      <Private>false</Private>
    </ProjectReference>
  </ItemGroup>

  <ItemGroup>
    <!-- INTEROP REFERENCES — vary by app, see Step 2 -->
    <Reference Include="Vendor.App.Interop">
      <HintPath>$(VendorInstallDir)\Interop\Vendor.App.Interop.dll</HintPath>
      <EmbedInteropTypes>false</EmbedInteropTypes>
      <Private>true</Private>
    </Reference>
  </ItemGroup>

  <Target Name="CopyToPluginsRoot" AfterTargets="Build">
    <PropertyGroup>
      <PluginsRoot>$(MSBuildThisFileDirectory)..\..\..\plugins\<AppName></PluginsRoot>
    </PropertyGroup>
    <ItemGroup>
      <PluginDeploy Include="$(OutputPath)*.dll" />
    </ItemGroup>
    <MakeDir Directories="$(PluginsRoot)" />
    <Copy SourceFiles="@(PluginDeploy)" DestinationFolder="$(PluginsRoot)" SkipUnchangedFiles="true" />
  </Target>
</Project>
```

If the interop path is machine-specific, **import `Common.Paths.props`** at
the top and use the resolution chain (see `LLM/paths.md` § "Per-plugin
csproj pattern"). Add the required interop DLLs to `@(RequiredInteropFile)`
so build-time validation catches missing installs with `error COMBRIDGE001`.

## Step 4 — Write the plugin class

Skeleton for any plugin (`<AppName>Plugin.cs`):

```csharp
using ComBridge.Core;
using Microsoft.CodeAnalysis;
using App = global::Vendor.App.Interop;   // ALIAS the interop namespace. Critical: see Pitfall 1 below.

namespace ComBridge.Plugins.<AppName>;

public sealed class <App>Globals
{
    // Script-facing fields. Cast to the DISPINTERFACE (_Application), not the co-class.
    // The _-prefixed type matches what ROT-fetched COM objects return; co-class casts
    // throw E_NOINTERFACE. See claude_code/lesson_20260521_excel_dispinterface_vs_coclass_cast.md.
    public App._Application app { get; }
    public App.Document? doc { get; }   // null-safe — the app may have no document open

    internal <App>Globals(App._Application a)
    {
        app = a;
        try { doc = a.ActiveDocument; } catch { doc = null; }
    }
}

public sealed class <App>Plugin : IComBridgePlugin
{
    public string Name => "<lowercase-name>";   // CLI name: combridge <Name> ...
    public string Description => "<short user-facing description>. Globals: app, doc.";
    public string[] ProgIds => new[] { "<Vendor>.<App>.Application" };
    public bool AllowCreateNew => true;   // false for heavy apps (CAD-class) — never silent-launch
    public Type GlobalsType => typeof(<App>Globals);

    public object CreateGlobals(object comRoot) => new <App>Globals((App._Application)comRoot);

    public IEnumerable<MetadataReference> ScriptReferences
    {
        get
        {
            var here = Path.GetDirectoryName(typeof(<App>Plugin).Assembly.Location)!;
            foreach (var dll in Directory.EnumerateFiles(here, "Vendor.App.Interop*.dll"))
                yield return MetadataReference.CreateFromFile(dll);
        }
    }

    public IEnumerable<string> ScriptUsings => new[] { "Vendor.App.Interop" };

    public IEnumerable<IBridgeCommand> Commands => new IBridgeCommand[]
    {
        new <App>InfoCommand(),   // minimum useful starting command
    };

    // Pattern A (document app): match file extensions in ROT.
    public IEnumerable<string> RotMonikerPatterns => new[]
    {
        @"\.(ext1|ext2|ext3)$",
    };

    // Pattern A: ascend bound Document RCW to its parent Application.
    public object? TryExtractRoot(object monikerBound)
    {
        try
        {
            dynamic d = monikerBound;
            object? app = d.Application;
            return app;
        }
        catch
        {
            return null;
        }
    }

    public (int? Pid, string? Title) DescribeInstance(object comRoot)
    {
        try
        {
            var app = (App._Application)comRoot;
            int? pid = null;
            try
            {
                // HWND retrieval — varies per app, see Step 5.
                var hwnd = (IntPtr)app.HWND;       // OR app.ActiveWindow.Hwnd, OR Process.GetProcessesByName(...)
                pid = SessionPicker.PidFromHwnd(hwnd);
            }
            catch { }

            string? title = null;
            try { title = app.ActiveDocument?.Name; } catch { }
            return (pid, title);
        }
        catch
        {
            return (null, null);
        }
    }
}

internal sealed class <App>InfoCommand : IBridgeCommand
{
    public string Name => "info";
    public string Usage => "info   (prints app version + active document)";

    public Task<int> RunAsync(object comRoot, string[] args, TextWriter output)
    {
        var app = (App._Application)comRoot;
        output.WriteLine($"<App> version: {app.Version}");
        try { output.WriteLine($"ActiveDocument: {app.ActiveDocument?.Name ?? "(none)"}"); }
        catch { output.WriteLine("ActiveDocument: (unavailable)"); }
        return Task.FromResult(0);
    }
}
```

## Step 5 — Wire up DescribeInstance correctly

HWND retrieval is the most app-specific piece. Use this table for known
apps, then probe for new ones:

| App | HWND access | Notes |
|---|---|---|
| Excel | `xlApp.Hwnd` (int) | lowercase `Hwnd` |
| Word | `wdApp.ActiveWindow.Hwnd` (int) | no HWND on Application directly; ActiveWindow can be null |
| PowerPoint | `pptApp.HWND` (int) | UPPERCASE `HWND` |
| Outlook | not directly exposed | use `Process.GetProcessesByName("OUTLOOK")[0].Id` — single-instance per user |
| AutoCAD (full) | `acadApp.HWND` (uppercase int) OR `acadApp.MainWindow.HWND` | varies by version |
| Inventor | `invApp.MainFrameHWND` (int) | non-standard name |
| SolidWorks | `((IFrame)swApp.Frame()).GetHWndx64()` (long) | use `GetHWndx64`, not `GetHWnd` (truncates on x64) |
| Visio | `vsoApp.WindowHandle32` (int) | |
| Photoshop / Illustrator | not directly exposed; rare automation use | use Process.GetProcessesByName fallback |

For an unknown app, write a one-line .csx in the partially-built plugin:

```csharp
Console.WriteLine(string.Join(", ",
    typeof(_Application).GetProperties().Select(p => p.Name)));
```

Look for properties named `HWND`, `Hwnd`, `MainWindow*`, `ActiveWindow*`,
`hWnd`. Try each.

Wrap every COM call in DescribeInstance in try/catch and return `(null, null)`
on any failure. SessionPicker's sidecar filter will drop dead bindings;
your job is to populate PID + Title when possible.

## Step 6 — Decide AllowCreateNew

| App class | `AllowCreateNew` | Why |
|---|---|---|
| Heavy CAD/engineering | `false` | SolidWorks-style: ~30s cold start, 1+ GB RAM. Silently launching is hostile. User should already have it open. |
| Lightweight Office | `true` | Excel, Word, PowerPoint start fast; users expect launch-on-demand. |
| Service apps (Outlook) | `true` | MAPI single-instance; if it's not running, scripted automation can plausibly launch it. |
| Reader-only / viewers | `false` (usually) | Often can't do meaningful automation without a doc loaded; let user open it first. |

This is a judgment call. When in doubt, prefer `false` (more conservative;
user sees no surprise process spawn).

## Step 7 — Add to the solution

Edit `combridge.sln`:

1. Add a `Project(...) = "ComBridge.Plugins.<App>"` line in the project list.
   Use a fresh GUID for the project (sequential is fine: `{A0000008-0000-0000-0000-000000000008}`).
2. Add four lines in `GlobalSection(ProjectConfigurationPlatforms) = postSolution`:
   ```
   {GUID}.Debug|Any CPU.ActiveCfg = Debug|Any CPU
   {GUID}.Debug|Any CPU.Build.0 = Debug|Any CPU
   {GUID}.Release|Any CPU.ActiveCfg = Release|Any CPU
   {GUID}.Release|Any CPU.Build.0 = Release|Any CPU
   ```

## Step 8 — Build and verify

```powershell
cd <repo-root>
dotnet build -c Release
```

If the build fails:
- Read the error message
- Cross-reference `LLM/troubleshooting.md` § "Build-time errors"
- Common: missing reference, namespace shadowing (alias collides with C# namespace), `<EmbedInteropTypes>false</EmbedInteropTypes>` missing, transitive PIA not copied

After build success, verify the plugin loads:

```powershell
.\combridge.exe list-plugins
# Should show your new plugin

.\combridge.exe <name> list-commands
# Should show: run-script (built-in) + list-sessions (built-in) + your custom commands

.\combridge.exe <name> list-sessions -
# If the target app is running: shows session(s) with pid + title
# If not running: "(no running <name> sessions in the ROT)"

.\combridge.exe <name> info -
# Exercises CreateGlobals + your IBridgeCommand.RunAsync
```

Then exercise selectors:

```powershell
.\combridge.exe <name> --session 1 info -
.\combridge.exe <name> --session pid:<PID> info -
.\combridge.exe <name> --session "<title-substring>" info -
```

And finally run-script with a `dynamic` test:

```csharp
// test.csx
dynamic d = app;
Console.WriteLine($"version: {d.Version}");
```

```powershell
.\combridge.exe <name> run-script test.csx -
```

If any of these fail, the plugin isn't done. See troubleshooting.

## Worked examples for shippable-by-LLM plugins

These are the apps you'd plausibly want a plugin for, with the exact
shape each plugin should take. We don't ship these because we don't have
the apps installed, but an LLM can construct them from these specs.

### AutoCAD (full version, NOT LT)

Pattern: **A** (document-based, file monikers) + **D** (version-suffixed ProgID).

```csharp
public string[] ProgIds => new[]
{
    "AutoCAD.Application",        // latest installed full AutoCAD
    "AutoCAD.Application.26",     // AutoCAD 2026 specifically
    "AutoCAD.Application.25",     // AutoCAD 2025
    "AutoCAD.Application.24",     // AutoCAD 2024
};
public bool AllowCreateNew => false;   // ~45s cold start; never silent-launch

public IEnumerable<string> RotMonikerPatterns => new[]
{
    @"\.(dwg|dxf|dwt|dws)$",     // drawings, templates, standards
};

public object? TryExtractRoot(object monikerBound)
{
    try { dynamic d = monikerBound; return (object)d.Application; }
    catch { return null; }
}

// HWND via app.HWND or app.MainWindow.HWND — varies by version.
// Title via app.ActiveDocument.Name.
```

Interop locations:
- `C:\Program Files\Autodesk\AutoCAD <ver>\AcCoreMgd.dll`
- `C:\Program Files\Common Files\Autodesk Shared\Autodesk.AutoCAD.Interop.dll`
- `C:\Program Files\Common Files\Autodesk Shared\Autodesk.AutoCAD.Interop.Common.dll`

Use `Common.Paths.props` with registry probe:
- `HKLM\SOFTWARE\Autodesk\AutoCAD\<ver>\<lang>\AcadLocation`

Globals: `acadApp`, `acadDoc`, `modelSpace`, `paperSpace`.

**Compatibility note:** does NOT work with AutoCAD LT (any version — LT
strips automation entirely). Often works with BricsCAD (a third-party
AutoCAD clone with intentionally-compatible COM) and AutoCAD-based
verticals (Civil 3D, Architecture, Mechanical, MEP, Electrical, Plant 3D)
with no plugin changes.

### Inventor (Autodesk 3D CAD)

Pattern: **A** (file monikers) — same as AutoCAD but different extensions.

```csharp
public string[] ProgIds => new[] { "Inventor.Application" };
public bool AllowCreateNew => false;   // CAD-class heavy app

public IEnumerable<string> RotMonikerPatterns => new[]
{
    @"\.(ipt|iam|idw|ipn|dwg)$",   // parts, assemblies, drawings, presentations
};

// HWND via app.MainFrameHWND (note non-standard naming).
// Title via app.ActiveDocument.DisplayName.
```

Interop: `C:\Program Files\Autodesk\Inventor <ver>\Bin\Autodesk.Inventor.Interop.dll`.
Registry probe: `HKLM\SOFTWARE\Autodesk\Inventor\<ver>\InstallPath`.

Globals: `invApp`, `invDoc`, `invAssy`, `invPart`, `invDrawing`.

### Adobe Acrobat Pro (PDF authoring)

Pattern: **C** (single-instance class-moniker). Acrobat doesn't register
per-PDF monikers in the ROT the way Office apps do.

```csharp
public string[] ProgIds => new[] { "AcroExch.App" };   // note: not "Acrobat.Application"
public bool AllowCreateNew => true;

public IEnumerable<string> RotMonikerPatterns => Array.Empty<string>();
// Rely on TryCoGetActiveObject only.
```

Interop: Acrobat ships `Acrobat.tlb`; you may need to `tlbimp` it or use
a pre-built `Interop.Acrobat.dll` from the Acrobat SDK. Multi-instance
Acrobat is unusual; single MAPI-style session works fine.

Globals: `acroApp` (`CAcroApp`), `acroAVDoc` (active PDF, `CAcroAVDoc`).

### Photoshop / Illustrator

Pattern: **C** (single-instance class-moniker). Same shape as Acrobat.

```csharp
// Photoshop
public string[] ProgIds => new[] { "Photoshop.Application" };
// Illustrator
public string[] ProgIds => new[] { "Illustrator.Application" };

public IEnumerable<string> RotMonikerPatterns => Array.Empty<string>();
```

Both expose a rich JavaScript-style scripting interface via COM. The
typed interop is thinner than Office's — be prepared to use `dynamic`
for many operations.

### BricsCAD

Pattern: **A or D**, intentionally AutoCAD-compatible.

```csharp
public string[] ProgIds => new[]
{
    "BricscadApp.AcadApplication",   // BricsCAD-native ProgID
    "AutoCAD.Application",            // BricsCAD ALSO registers under AutoCAD's ProgID for compat
};
public IEnumerable<string> RotMonikerPatterns => new[] { @"\.(dwg|dxf)$" };
```

A combined `autocad-or-bricscad` plugin is feasible since the COM surface
is intentionally parallel. Globals can stay named `acadApp` etc.

### Visio

Pattern: **A** (file monikers) + **D** (version-suffixed ProgID).

```csharp
public string[] ProgIds => new[]
{
    "Visio.Application",
    "Visio.Application.16",   // Visio 2016/365
};
public IEnumerable<string> RotMonikerPatterns => new[]
{
    @"\.(vsdx|vsdm|vstx|vstm|vssx|vssm|vsd)$",
};
// HWND via app.WindowHandle32 (note non-standard property name)
```

Interop: ships with Office, GAC location like Word/PowerPoint.
Globals: `vsoApp`, `vsoDoc`, `vsoPage`.

## Common pitfalls (cross-references to lessons)

For each, the symptom from the wild + where to read more:

| Pitfall | Symptom | Read |
|---|---|---|
| Casting RCW to co-class instead of dispinterface | `InvalidCastException: cannot cast … to _Application … QueryInterface for IID '{…}' failed (E_NOINTERFACE)` | `LLM/troubleshooting.md` § "Cast errors"; `C:\personal_rag\claude_code\lesson_20260521_excel_dispinterface_vs_coclass_cast.md` |
| Missing `office.dll` for Office plugins | `FileNotFoundException: office, Version=15.0.0.0, PublicKeyToken=71e9bce111e9429c` | Same lesson; add `<Reference Include="office">` HintPath to GAC |
| Namespace shadowing (`using Excel = ...Excel` collides with plugin's `Excel` namespace) | `CS0234: type 'Application' does not exist in namespace 'ComBridge.Plugins.Excel'` | `LLM/build.md`; alias to a non-clashing name (`Xl`, `Wd`, `Pp`, `Ol`) |
| ROT walk finds nothing | `list-sessions` is empty when app is running | Almost certainly wrong `RotMonikerPatterns`. Dump the ROT (see Step 1) and look at actual display names. |
| Found a binding but it's the wrong type | Cast in DescribeInstance throws | Your `RotMonikerPatterns` matched some other COM object's moniker. Tighten the regex with anchors `^…$`. |
| dynamic in scripts fails | `CS0656: Missing compiler required member 'Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo.Create'` | Already handled — ScriptHost includes `Microsoft.CSharp` in default refs. If a script still fails with this, your plugin is missing `using Microsoft.CSharp` in `ScriptUsings`. |
| Zombie sidecar processes appear in list-sessions | `(no info)` rows showing | Already handled — SessionPicker drops entries where both PID and Title are null. If you see a "(no info)" row, your DescribeInstance is returning a valid PID OR a valid title for a transient COM object — investigate why. |
| Plugin loads but globals are null at runtime | `NullReferenceException` on first script line | Cast in `CreateGlobals` failed silently due to ALC mismatch. Verify `InteractiveAssemblyLoader.RegisterDependency` includes your plugin assembly — handled in `ScriptHost.cs`. |

## Verification checklist (do not skip)

Tick all of these before declaring the plugin done:

- [ ] `dotnet build -c Release` succeeds with 0 errors, 0 warnings
- [ ] `plugins/<App>/` contains the plugin DLL + interop DLLs
- [ ] `combridge list-plugins` shows the new plugin
- [ ] `combridge <name> list-commands` shows `run-script`, `list-sessions`, and any custom commands
- [ ] If the app is running on this machine: `combridge <name> list-sessions -` shows at least one entry with PID + title
- [ ] If the app is running: `combridge <name> info -` returns useful output (version + active doc)
- [ ] Selectors all work:
  - [ ] `--session 1 info -` (this is the MRU session per Z-order)
  - [ ] `--session last info -` (synonym keyword — same as `--session 1`)
  - [ ] `--session pid:<PID> info -`
  - [ ] `--session "<title>" info -`
- [ ] **MRU sanity check**: with 2+ instances running, click in window B then
      run `combridge <name> active-doc -` from a terminal. Output should
      reference B (not A). If it always picks the same instance regardless of
      focus, `DescribeInstance` likely isn't returning a PID, which disables
      Z-order ranking — fix that first.
- [ ] A minimal `run-script` that does `Console.WriteLine(app.Version)` succeeds
- [ ] If `dynamic` is used in a script, it compiles and runs (validates `Microsoft.CSharp` ref + ALC binding)
- [ ] After the test app is closed, `list-sessions` reports empty cleanly (no exception)
- [ ] `LLM/plugins.md` updated with a new section (member tables, RotMonikerPatterns/TryExtractRoot/DescribeInstance specifics, any gotchas)
- [ ] `LLM/symbols.md` updated with type rows + deployment-path row
- [ ] `LLM/README.md` defaults reference table updated
- [ ] If the plugin shipped any new architectural learning (a new ROT pattern, an undocumented HWND property, etc.) — write a lesson to `C:\personal_rag\claude_code\` and update the indexes

## What NOT to do

- Don't add the plugin DLL to git if it can't be built from source on a checkout. Plugins must build from source.
- Don't reference per-machine paths in code (use `Common.Paths.props` chain).
- Don't disable `Common.Paths.props` validation. If `RequiredInteropFile` is missing on a developer's machine, fail the build loudly with COMBRIDGE001.
- Don't write a plugin for an app that doesn't expose COM. AutoCAD LT is the prototypical example — there's nothing for combridge to bridge to.
- Don't introduce per-plugin selector grammars or new `--flags`. The selector grammar in `LLM/cli.md` is the cross-plugin contract.
- Don't break the stability tier promises in `LLM/api.md`. If a Core API needs to change, bump the tier or version, don't quietly mutate.
