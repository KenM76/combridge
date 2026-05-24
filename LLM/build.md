# Build & Deploy — Pitfalls and Recipes

## Layout invariants

- `Directory.Build.props` at repo root sets `TargetFramework=net10.0-windows`, `Nullable=enable`, `ImplicitUsings=enable`. All projects inherit it.
- ProjectReference from any plugin to `ComBridge.Core` MUST be `<Private>false</Private>` to avoid duplicating Core in plugin folders.
- Each plugin csproj has a `CopyToPluginsRoot` MSBuild target that runs `AfterTargets="Build"` and copies `$(OutputPath)*.dll` to `$(MSBuildThisFileDirectory)..\..\..\plugins\<Name>\`.
- The host exe must be staged together with the `plugins/` folder. PluginLoader resolves `<exeDir>/plugins/<Name>/`.

## Pitfalls table

| Symptom | Cause | Fix |
|---|---|---|
| `CS0246: type or namespace 'IRunningObjectTable' not found` | `IRunningObjectTable`, `IBindCtx`, `IMoniker` live in `System.Runtime.InteropServices.ComTypes`, not `.InteropServices` | `using System.Runtime.InteropServices.ComTypes;` |
| `CS0234: type 'Application' does not exist in namespace 'ComBridge.Plugins.Excel'` | Plugin namespace ends in `.Excel`; `using Excel = Microsoft.Office.Interop.Excel` resolves `Excel.Application` against the local namespace, not the alias | Alias to a non-clashing name, e.g. `using Xl = global::Microsoft.Office.Interop.Excel;` |
| Plugin loads but `run-script` fails to compile (interop types unresolved) | NuGet PIA package defaults to `EmbedInteropTypes=true`, inlines types, doesn't copy DLL to output | Set `<EmbedInteropTypes>false</EmbedInteropTypes>` on the `PackageReference` |
| Interop DLL still not in plugin output folder after EmbedInteropTypes fix | NuGet transitive deps not copied by default | Add `<CopyLocalLockFileAssemblies>true</CopyLocalLockFileAssemblies>` to `<PropertyGroup>` |
| Excel plugin folder grew by ~10 MB | `CopyLocalLockFileAssemblies=true` also copies Roslyn into the plugin folder | Acceptable — `PluginLoadContext.Load` prefers default-context assemblies, so duplicates aren't loaded twice |
| Plugins aren't discovered at runtime | Host exe not staged next to a `plugins/` directory; or plugin DLL name doesn't start with `ComBridge.Plugins.` | Either copy exe + plugins folder together, or rename plugin DLL to match `ComBridge.Plugins.<dirName>.dll` |
| `Marshal.GetActiveObject` not found at compile time | Removed in .NET (Core/5+); only in .NET Framework | Use `RotHelper.TryGetActiveObject` (this project) — implements via `ole32!GetRunningObjectTable` + `CreateBindCtx` |
| `--session pid:NNNN` never matches | Plugin doesn't override `DescribeInstance` → `Pid` always null | Implement `DescribeInstance` returning `(SessionPicker.PidFromHwnd((IntPtr)appHwnd), title)` |
| `IFrame` cast returns null on a SW session that's mid-startup | `as`-on-RCW can silently null on some COM objects | Always wrap `DescribeInstance` body in try/catch and return `(null, null)` on failure — instance still gets an index |
| `error COMBRIDGE001: could not locate ... interop assemblies` | Resolved interop path doesn't contain the DLLs (SW not installed, non-standard location, wrong env var) | Read the error's resolution-chain block; set `paths.props` / env var / `/p:` so the listed files exist. Full spec: `LLM/paths.md` |
| `paths.props` edits have no effect | `Common.Paths.props` imported AFTER the plugin's fallback `<PropertyGroup>`, so env/registry/default already won | Import `..\Common.Paths.props` at the TOP of the .csproj, before the resolution PropertyGroup |
| `paths.props` from another plugin bleeds in | Used `$(MSBuildThisFileDirectory)` (Common.Paths.props's dir) instead of project dir | Resolution uses `$(MSBuildProjectDirectory)\paths.props` — the importing project's dir. Don't change this |
| Registry probe property is empty on a machine that has the app | Wrong hive view (32 vs 64) or value name | Use `RegistryView.Registry64`; verify exact value name (see `LLM/paths.md` § "Registry probe — known keys") |
| `paths.props` got committed to git | `.gitignore` not updated | `.gitignore` excludes `paths.props` (the live file); only `paths.props.example` is tracked |

## csproj boilerplate (copy verbatim)

### ComBridge.Core (already in repo)

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Library</OutputType>
    <RootNamespace>ComBridge.Core</RootNamespace>
    <AssemblyName>ComBridge.Core</AssemblyName>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="Microsoft.CodeAnalysis.CSharp.Scripting" Version="4.13.0" />
  </ItemGroup>
</Project>
```

### Plugin using NuGet PIA

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Library</OutputType>
    <RootNamespace>ComBridge.Plugins.MyApp</RootNamespace>
    <AssemblyName>ComBridge.Plugins.MyApp</AssemblyName>
    <PlatformTarget>x64</PlatformTarget>
    <CopyLocalLockFileAssemblies>true</CopyLocalLockFileAssemblies>
  </PropertyGroup>
  <ItemGroup>
    <ProjectReference Include="..\..\ComBridge.Core\ComBridge.Core.csproj">
      <Private>false</Private>
    </ProjectReference>
    <PackageReference Include="Microsoft.Office.Interop.MyApp" Version="...">
      <EmbedInteropTypes>false</EmbedInteropTypes>
    </PackageReference>
  </ItemGroup>
  <Target Name="CopyToPluginsRoot" AfterTargets="Build">
    <PropertyGroup>
      <PluginsRoot>$(MSBuildThisFileDirectory)..\..\..\plugins\MyApp</PluginsRoot>
    </PropertyGroup>
    <ItemGroup>
      <PluginDeploy Include="$(OutputPath)*.dll" />
    </ItemGroup>
    <MakeDir Directories="$(PluginsRoot)" />
    <Copy SourceFiles="@(PluginDeploy)" DestinationFolder="$(PluginsRoot)" SkipUnchangedFiles="true" />
  </Target>
</Project>
```

### Plugin using non-NuGet interop (HintPath)

```xml
<ItemGroup>
  <ProjectReference Include="..\..\ComBridge.Core\ComBridge.Core.csproj">
    <Private>false</Private>
  </ProjectReference>
  <Reference Include="MyApp.Interop">
    <!-- HintPath is machine-specific; adjust to wherever the app's interop ships. -->
    <HintPath>C:\Program Files\MyApp\API\MyApp.Interop.dll</HintPath>
    <EmbedInteropTypes>false</EmbedInteropTypes>
    <Private>true</Private>
  </Reference>
</ItemGroup>
<!-- No CopyLocalLockFileAssemblies needed for <Reference> path. -->
```

## Build commands

```powershell
cd <repo-root>                                                     # dir containing combridge.sln
dotnet build -c Release                                            # full solution
dotnet build src/plugins/ComBridge.Plugins.MyApp -c Release        # one plugin
```

After build:
- Host exe: `src/ComBridge.Cli/bin/Release/net10.0-windows/combridge.exe`
- Plugins: `<repo-root>/plugins/<Name>/` (auto-deployed by CopyToPluginsRoot)
- Stage for use: copy `combridge.exe` + all `*.dll`/`*.json` from the Cli output dir to the same place as the `plugins/` folder.

## Solution registration (combridge.sln)

For a new plugin `MyApp` with GUID `A0000005-...-005` (any unique GUID):

```
Project("{9A19103F-16F7-4668-BE54-9A1E7A4F7556}") = "ComBridge.Plugins.MyApp", "src\plugins\ComBridge.Plugins.MyApp\ComBridge.Plugins.MyApp.csproj", "{A0000005-0000-0000-0000-000000000005}"
EndProject
```

Plus 4 lines in `GlobalSection(ProjectConfigurationPlatforms) = postSolution`:

```
{A0000005-0000-0000-0000-000000000005}.Debug|Any CPU.ActiveCfg = Debug|Any CPU
{A0000005-0000-0000-0000-000000000005}.Debug|Any CPU.Build.0 = Debug|Any CPU
{A0000005-0000-0000-0000-000000000005}.Release|Any CPU.ActiveCfg = Release|Any CPU
{A0000005-0000-0000-0000-000000000005}.Release|Any CPU.Build.0 = Release|Any CPU
```

## Verification checklist

After any change to Core or a plugin, run:

```powershell
dotnet build -c Release
ls .\plugins\<PluginName>\                          # plugin DLL + interop must be there (run from repo root)
.\combridge.exe list-plugins                       # plugin must appear
.\combridge.exe <plugin> list-commands             # built-ins (run-script, list-sessions) + custom
.\combridge.exe <plugin> list-sessions -           # check if app is running
```

If `list-sessions` shows entries with `(no info)`, `DescribeInstance` is throwing or returning `(null, null)`. Add try/catch logging to debug.
