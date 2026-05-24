# Machine-Specific Path Resolution

Plugins that reference interop assemblies by file path (not NuGet) need a
per-machine path. The resolution is implemented in MSBuild, shared via
`src/plugins/Common.Paths.props`, and is cross-platform clean (only the
registry layer is Windows-gated).

## Resolution chain (first non-empty wins)

| # | Layer | Mechanism | Cross-platform | Set by |
|---|---|---|---|---|
| 1 | Explicit property | `dotnet build /p:<Prop>=<value>` (MSBuild global property — cannot be overwritten by `<PropertyGroup>`) | ✓ | CI / one-off |
| 2 | `paths.props` | `<plugin-dir>/paths.props`, imported by `Common.Paths.props` via `$(MSBuildProjectDirectory)\paths.props` | ✓ | user copies from `paths.props.example` |
| 3 | Environment variable | `$([System.Environment]::GetEnvironmentVariable('<VAR>'))` in plugin csproj | ✓ | shell / system env |
| 4 | Registry auto-probe | `$([MSBuild]::GetRegistryValueFromView(...))` gated `Condition="'$(OS)' == 'Windows_NT'"` | Windows only | app installer (automatic) |
| 5 | Built-in default | hardcoded canonical path in plugin csproj | ✓ (OS-conditioned defaults possible) | repo |
| 6 | Validation | `ValidateRequiredInteropFiles` target → `error COMBRIDGE001` | ✓ | (failure path) |

Why this order: explicit override beats persisted local config beats
shell env beats auto-detect beats default. Each csproj assignment is
gated `Condition="'$(Prop)' == ''"`, so layers 1–2 (already set before
the plugin's PropertyGroup runs) automatically short-circuit 3–5.

## Mechanism

`src/plugins/Common.Paths.props`:

```xml
<Import Project="$(MSBuildProjectDirectory)\paths.props"
        Condition="Exists('$(MSBuildProjectDirectory)\paths.props')" />

<Target Name="ValidateRequiredInteropFiles"
        BeforeTargets="ResolveAssemblyReferences;CoreCompile"
        Condition="'@(RequiredInteropFile)' != ''">
  <ItemGroup>
    <_MissingInteropFile Include="@(RequiredInteropFile)"
                         Condition="!Exists('%(RequiredInteropFile.Identity)')" />
  </ItemGroup>
  <Error Condition="'@(_MissingInteropFile)' != ''" Code="COMBRIDGE001"
         Text="... Missing file(s): @(_MissingInteropFile,...) ... $(RequiredInteropDiagnostic) ..." />
</Target>
```

Key facts:
- `$(MSBuildProjectDirectory)` = directory of the **importing .csproj**, NOT of `Common.Paths.props`. So each plugin's `paths.props` is local to that plugin.
- `Common.Paths.props` must be `<Import>`ed BEFORE the plugin's fallback `<PropertyGroup>`, else a `paths.props` value won't pre-empt the env/registry/default assignments.
- Validation target runs `BeforeTargets="ResolveAssemblyReferences;CoreCompile"` → fails before the cryptic "metadata file not found".
- Plugin contract with the shared file:
  - `@(RequiredInteropFile)` — item list of absolute file paths that MUST exist.
  - `$(RequiredInteropDiagnostic)` — string appended to the error; plugin sets it to show resolved value + env/registry names.

## Per-plugin csproj pattern

```xml
<Import Project="..\Common.Paths.props" />
<PropertyGroup>
  <!-- Layer 3 -->
  <FooApiDir Condition="'$(FooApiDir)' == ''">$([System.Environment]::GetEnvironmentVariable('FOO_API_DIR'))</FooApiDir>
  <!-- Layer 4 (Windows only) -->
  <_FooReg Condition="'$(FooApiDir)' == '' and '$(OS)' == 'Windows_NT'">$([MSBuild]::GetRegistryValueFromView('HKEY_LOCAL_MACHINE\SOFTWARE\Foo\Setup', 'InstallDir', null, RegistryView.Registry64))</_FooReg>
  <FooApiDir Condition="'$(FooApiDir)' == '' and '$(_FooReg)' != ''">$([System.IO.Path]::Combine('$(_FooReg)', 'api'))</FooApiDir>
  <!-- Layer 5 -->
  <FooApiDir Condition="'$(FooApiDir)' == ''">C:\Program Files\Foo\api</FooApiDir>
  <RequiredInteropDiagnostic>FooApiDir resolved to: '$(FooApiDir)'%0a(env var: FOO_API_DIR | registry: HKLM\SOFTWARE\Foo\Setup\InstallDir)%0a</RequiredInteropDiagnostic>
</PropertyGroup>
<ItemGroup>
  <Reference Include="Foo.Interop"><HintPath>$(FooApiDir)\Foo.Interop.dll</HintPath>
    <EmbedInteropTypes>false</EmbedInteropTypes><Private>true</Private></Reference>
</ItemGroup>
<ItemGroup>
  <RequiredInteropFile Include="$(FooApiDir)\Foo.Interop.dll" />
</ItemGroup>
```

`%0a` in `RequiredInteropDiagnostic` is an MSBuild newline escape; the console logger renders each as its own prefixed line (expected, still readable).

## Registry probe — known keys

| App | Hive\Key | Value | Yields | api subfolder |
|---|---|---|---|---|
| SolidWorks | `HKLM\SOFTWARE\SolidWorks\Setup` | `Solidworks Folder` | install root (trailing `\`) | `api\redist` |
| AutoCAD | `HKLM\SOFTWARE\Autodesk\AutoCAD\<ver>\<lang>` | `AcadLocation` | install dir | (interop in install dir) |
| Inventor | `HKLM\SOFTWARE\Autodesk\Inventor\<ver>` | `InstallPath` | install dir | `Bin` |

Use `RegistryView.Registry64` (64-bit apps). `GetRegistryValueFromView` returns empty string (not error) when the key/value is absent → chain falls through cleanly. Normalize trailing slashes with `$([System.IO.Path]::Combine(...))`.

## Plugins NOT using this

A plugin sourcing interop from a **NuGet PIA package** needs none of this —
NuGet resolves the DLL. The Excel plugin is the reference example: it has
no `paths.props`, no `Common.Paths.props` import, no `RequiredInteropFile`.
See `LLM/build.md` for the NuGet PIA csproj requirements
(`EmbedInteropTypes=false` + `CopyLocalLockFileAssemblies=true`).

## Files

| File | Committed? | Role |
|---|---|---|
| `src/plugins/Common.Paths.props` | yes | shared import + validation target |
| `src/plugins/<Plugin>/paths.props.example` | yes | documented template (layer 2) |
| `src/plugins/<Plugin>/paths.props` | NO (gitignored) | live per-machine override (layer 2) |

## External consumers (library-mode CLIs)

Third-party tools that take a HintPath dependency on `ComBridge.Core` (see
`LLM/consuming.md`) reuse this same path-resolution chain. The mechanism
is identical to a plugin — only the import path changes:

```xml
<!-- in the external consumer's .csproj -->
<PropertyGroup>
  <ComBridgeRoot>$([System.Environment]::GetEnvironmentVariable('COMBRIDGE_ROOT'))</ComBridgeRoot>
  <ComBridgeRoot Condition="'$(ComBridgeRoot)' == ''">D:\Dev\combridge</ComBridgeRoot>
</PropertyGroup>

<Import Project="$(ComBridgeRoot)\src\plugins\Common.Paths.props" />

<!-- Then declare your own interop paths and RequiredInteropFile items
     exactly as a plugin csproj would (see "Per-plugin csproj pattern" above). -->
```

Key facts for external consumers:
- `$(MSBuildProjectDirectory)` inside `Common.Paths.props` resolves to the **importing project's** directory, so a `paths.props` next to the consumer's `.csproj` is the one that gets imported. No cross-contamination from the combridge tree's own plugins.
- The validation target uses error code `COMBRIDGE001` regardless of whether it fires from inside or outside the combridge tree.
- `$(ComBridgeRoot)` defaults are the consumer's responsibility — combridge does not provide one. The pattern above is the canonical three-layer chain (`/p:` > env var > default).
- Once a NuGet package or `releases/<version>/` convention exists (planned, see `CONSUMING_CORE.md` § "Versioning"), the preferred import path will become `$(ComBridgeRoot)\paths\Common.Paths.props` or a NuGet content-file. Until then, use the in-tree path.

## Verification commands

```powershell
# default/registry resolution (no overrides)
dotnet build -c Release

# env-var layer, valid path  → Build succeeded
$env:SOLIDWORKS_API_REDIST="C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\api\redist"
dotnet build src/plugins/ComBridge.Plugins.SolidWorks/ComBridge.Plugins.SolidWorks.csproj -c Release

# any layer, bad path → error COMBRIDGE001 with full chain diagnostic, Build FAILED
$env:SOLIDWORKS_API_REDIST="Z:\nope"
dotnet build src/plugins/ComBridge.Plugins.SolidWorks/ComBridge.Plugins.SolidWorks.csproj -c Release
```

NOTE (Git Bash on Windows): MSYS mangles `/p:Prop=val` and a bare project
directory arg ("Only one project can be specified"). Use the env-var layer
to test, or pass the explicit `.csproj` path and `--property:Prop=val`.
