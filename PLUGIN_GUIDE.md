# Writing a combridge plugin

A plugin is a class library that:

1. References `ComBridge.Core`
2. References whatever interop assemblies its target app needs
3. Defines a globals type (the object exposed to user `.csx` scripts)
4. Implements `IComBridgePlugin`
5. Has a build target that copies its output into
   `<repo>/plugins/<PluginName>/` next to `combridge.exe`

## 1. Project file

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Library</OutputType>
    <PlatformTarget>x64</PlatformTarget>
    <!-- Required when using NuGet PIA packages, ignored otherwise. -->
    <CopyLocalLockFileAssemblies>true</CopyLocalLockFileAssemblies>
  </PropertyGroup>
  <ItemGroup>
    <ProjectReference Include="..\..\ComBridge.Core\ComBridge.Core.csproj">
      <Private>false</Private>  <!-- core lives next to the host, don't dup -->
    </ProjectReference>
    <!-- App interop, either: -->
    <PackageReference Include="Microsoft.Office.Interop.Word" Version="...">
      <EmbedInteropTypes>false</EmbedInteropTypes>  <!-- IMPORTANT: PIAs default to true -->
    </PackageReference>
    <!-- Or: -->
    <Reference Include="MyApp.Interop">
      <HintPath>C:\Program Files\MyApp\API\MyApp.Interop.dll</HintPath>
      <Private>true</Private>
    </Reference>
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

## 2. Globals type + plugin class

```csharp
using ComBridge.Core;
using Microsoft.CodeAnalysis;
using MyApp.Interop;

public sealed class MyAppGlobals
{
    public IApplication app { get; }
    public IDocument? doc { get; }
    internal MyAppGlobals(IApplication a) { app = a; doc = a.ActiveDocument; }
}

public sealed class MyAppPlugin : IComBridgePlugin
{
    public string Name => "myapp";
    public string Description => "MyApp automation. Globals: app, doc.";
    public string[] ProgIds => new[] { "MyApp.Application" };
    public bool AllowCreateNew => true;          // false = attach only
    public Type GlobalsType => typeof(MyAppGlobals);
    public object CreateGlobals(object root) => new MyAppGlobals((IApplication)root);

    public IEnumerable<MetadataReference> ScriptReferences
    {
        get
        {
            var here = Path.GetDirectoryName(typeof(MyAppPlugin).Assembly.Location)!;
            foreach (var dll in Directory.EnumerateFiles(here, "MyApp.Interop*.dll"))
                yield return MetadataReference.CreateFromFile(dll);
        }
    }

    public IEnumerable<string> ScriptUsings => new[] { "MyApp.Interop" };
    public IEnumerable<IBridgeCommand> Commands => new IBridgeCommand[] { /* optional */ };

    // STRONGLY RECOMMENDED: the PID enables MRU ordering (default attach behavior),
    // PID-based selection, and title-substring selection. Skip and you get unsorted
    // sessions with `(no info)` rows — the bridge still works but loses the "pick the
    // window the user was last on" guarantee. The cast + try/catch pattern below is
    // standard; see the SolidWorks/Excel/Word plugins for live examples.
    public (int? Pid, string? Title) DescribeInstance(object comRoot)
    {
        var app = (IApplication)comRoot;
        var hwnd = (IntPtr)app.HWND;                       // whatever HWND property the app exposes
        var pid  = SessionPicker.PidFromHwnd(hwnd);
        var title = app.ActiveDocument?.Name;
        return (pid, title);
    }
}
```

## 3. Plugin-specific commands (optional)

```csharp
internal sealed class StatusCommand : IBridgeCommand
{
    public string Name => "status";
    public string Usage => "status   (prints app + doc info)";
    public Task<int> RunAsync(object comRoot, string[] args, TextWriter output)
    {
        var app = (IApplication)comRoot;
        output.WriteLine(app.Version);
        return Task.FromResult(0);
    }
}
```

`run-script` and `list-sessions` are built in — you don't need to implement them.

End users can also drop their own `.csx` files into `plugins/<YourPlugin>/commands/`
and they'll auto-discover as named commands. Anything in your plugin's
globals (the type returned from `CreateGlobals`) is accessible from user
scripts. See `LLM/extending.md` for the convention; you as a plugin author
don't need to do anything special — the auto-discovery happens at the
host level. Your typed `IBridgeCommand` implementations win against
same-named scripts (so you control the canonical API surface).

## 3a. Machine-specific interop paths (HintPath plugins only)

Skip this entire section if your interop comes from a **NuGet PIA package**
(like the Excel plugin) — NuGet handles resolution and there's nothing to
configure.

If you reference interop by file path (`<Reference><HintPath>...`), don't
hardcode an absolute path. Use the shared resolution chain so the plugin
builds on any machine. Import the helper at the **top** of your csproj
(before the resolution `<PropertyGroup>`), resolve the path through the
fallback layers, and declare the files that must exist:

```xml
<Import Project="..\Common.Paths.props" />
<PropertyGroup>
  <!-- explicit /p: and paths.props (imported above) win automatically -->
  <FooApiDir Condition="'$(FooApiDir)' == ''">$([System.Environment]::GetEnvironmentVariable('FOO_API_DIR'))</FooApiDir>
  <_FooReg Condition="'$(FooApiDir)' == '' and '$(OS)' == 'Windows_NT'">$([MSBuild]::GetRegistryValueFromView('HKEY_LOCAL_MACHINE\SOFTWARE\Foo\Setup', 'InstallDir', null, RegistryView.Registry64))</_FooReg>
  <FooApiDir Condition="'$(FooApiDir)' == '' and '$(_FooReg)' != ''">$([System.IO.Path]::Combine('$(_FooReg)', 'api'))</FooApiDir>
  <FooApiDir Condition="'$(FooApiDir)' == ''">C:\Program Files\Foo\api</FooApiDir>
  <RequiredInteropDiagnostic>FooApiDir resolved to: '$(FooApiDir)'%0a(env var: FOO_API_DIR | registry: HKLM\SOFTWARE\Foo\Setup\InstallDir)%0a</RequiredInteropDiagnostic>
</PropertyGroup>
<ItemGroup>
  <RequiredInteropFile Include="$(FooApiDir)\Foo.Interop.dll" />
</ItemGroup>
```

Resolution order (first non-empty wins): `dotnet build /p:FooApiDir=...`
→ local `paths.props` → env var → Windows registry → built-in default.
If none yields a folder containing the `RequiredInteropFile`s, the build
fails fast with a readable `error COMBRIDGE001` listing every layer it
tried. Ship a `paths.props.example` next to the csproj documenting the
knobs (copy the SolidWorks plugin's). The live `paths.props` is gitignored.

Full reference: `LLM/paths.md`.

## 4. Conventions

- **Globals naming**: pick names that match the app's COM convention so users
  recognize them (`xlApp`, `swApp`, `wdApp`, `acadApp`). Don't invent generic
  names like `app` and `doc` if the app community already has standards.
- **AllowCreateNew**:
  - `true` for lightweight apps that are cheap to launch (Excel, Word)
  - `false` for heavy apps where a silent launch would surprise the user
    (SolidWorks, AutoCAD, Inventor)
- **Null-safety in globals**: ActiveDoc can be null, the active sheet might
  not be a Worksheet, etc. Guard in the globals constructor and let scripts
  null-check.
- **Casting COM RCWs**: when in doubt, prefer hard casts over `as` — the
  `as` operator silently returns null for some RCWs. Verify the underlying
  type first via the app's own API (e.g. `IModelDoc2.GetType()` in SW).
- **Namespace shadowing**: do NOT use the interop's top-level name as your
  plugin namespace's last segment. `namespace ComBridge.Plugins.Excel` plus
  `using Excel = Microsoft.Office.Interop.Excel` resolves `Excel.Application`
  against your namespace, not the alias. Use a different alias (`Xl`).
- **DescribeInstance**: wrap everything in try/catch and return `(null, null)`
  on failure — instances mid-startup may not have a reachable HWND yet.
