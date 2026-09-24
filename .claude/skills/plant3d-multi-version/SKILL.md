---
name: plant3d-multi-version
description: How to build and ship one Plant 3D plugin code base for several Plant 3D releases with different .NET runtimes — 2024 (.NET Framework 4.8), 2025/2026 (.NET 8), 2027+ (.NET 10) — via an SDK-style multi-targeted csproj, per-version SDK paths, conditional code, and an autoloader bundle. Use when converting the project file, adding a Plant 3D version, touching build/deployment, or writing code that must compile on both runtimes.
---

# One code base, several Plant 3D versions

| Plant 3D | AutoCAD base | .NET runtime | TFM | SDK folder (Drive) | Plant assembly version |
|---|---|---|---|---|---|
| 2024 | R24.3 | .NET Framework 4.8 | `net48` | `AP3D_2024_SDK` | 15.0 |
| 2025 | R25.0 | .NET 8 | `net8.0-windows` | not provided | 16.x |
| 2026 | R25.1 | .NET 8 | `net8.0-windows` | `AP3D_2026_SDK` | 17.0 |
| 2027+ | ? | .NET 10 (per user) | `net10.0-windows` | none yet | ? |

The 2024 and 2026 values in the table are verified from the SDK DLL metadata. The 2025 and 2027 rows are partly assumptions; confirm them when those SDKs are available.

## Key facts (verified)
- The Plant API the plugin uses is **the same** in 2024 and 2026. The unchanged sources compile for both targets
  (`tools/compile-check`). Porting is mostly about the project file and deployment, not the code.
- The 2026 SDK samples use SDK-style projects (`<Project Sdk="Microsoft.Net.Sdk">`) with `UseWindowsForms`,
  and reference `..\..\..\inc\AcCoreMgd.dll` and `..\..\..\inc-x64\PnP*.dll` with `Private=False`.
- A .NET 8 plugin cannot load in 2024, and a net48 plugin cannot load in 2025+. **You need one DLL per runtime.**

## Target project layout (not done yet; propose it to the user before converting)
Replace the old-style `WeldPropUtils.csproj` with an SDK-style project that multi-targets:

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFrameworks>net48;net8.0-windows</TargetFrameworks>   <!-- add net10.0-windows for 2027 -->
    <UseWindowsForms>true</UseWindowsForms>
    <PlatformTarget>x64</PlatformTarget>
    <GenerateAssemblyInfo>false</GenerateAssemblyInfo>            <!-- keep Properties\AssemblyInfo.cs -->
    <AppendTargetFrameworkToOutputPath>true</AppendTargetFrameworkToOutputPath>
  </PropertyGroup>
  <PropertyGroup Condition="'$(TargetFramework)'=='net48'">
    <LangVersion>7.3</LangVersion>
    <PlantSdk>$(AP3D_SDK_2024)</PlantSdk>
  </PropertyGroup>
  <PropertyGroup Condition="'$(TargetFramework)'=='net8.0-windows'">
    <PlantSdk>$(AP3D_SDK_2026)</PlantSdk>
  </PropertyGroup>
  <ItemGroup>
    <Reference Include="AcCoreMgd" HintPath="$(PlantSdk)\inc\AcCoreMgd.dll" Private="false" />
    <!-- … AcDbMgd, AcMgd from inc\ ; PnP*.dll from inc-x64\ … -->
  </ItemGroup>
</Project>
```
- Output: `bin\Release\net48\WeldPropUtils.dll` (for 2024) and `bin\Release\net8.0-windows\WeldPropUtils.dll` (for 2025/2026).
- Build against the **oldest** SDK of each runtime family, so the DLL uses nothing newer than that family has.
  Whether a DLL built against 2026 (17.0) loads in 2025 (16.x) is **untested**. If 2025 matters, get its SDK.
- Set the environment variables `AP3D_SDK_2024` and `AP3D_SDK_2026` on the build machine. The SDK folders on Drive
  already combine ObjectARX `inc\` (AcCoreMgd, …) with the Plant files in `inc-x64\`.
- Keep `LangVersion 7.3` on net48, so the shared code avoids newer C# syntax, or use features that compile for both.
  To use C# 8+ you need `LangVersion` `latest` on net48 too. That works for syntax-only features but not for ones
  that need runtime types (e.g. `Index`/`Range`, default interface methods).

## Code that differs between runtimes
- Try to have none. When unavoidable, use `#if NET8_0_OR_GREATER` / `#if NETFRAMEWORK`.
- These .NET Framework-only APIs don't exist on .NET 8/10: `System.Runtime.Remoting`, `AppDomain.CreateDomain`,
  `BinaryFormatter` (removed in .NET 9+), `System.Web`, WCF server, `Thread.Abort`.
- WinForms and WPF work on net8/net10 with `UseWindowsForms` / `UseWPF`. Keep a `MessageBox` out of loops anyway.
- Culture handling is the same on both. Always parse numbers with `CultureInfo.InvariantCulture`.

## Deployment: one autoloader bundle for every version
`WeldPropUtils.bundle\PackageContents.xml` with one `<Components>` block per runtime:
```xml
<Components Description="Plant 3D 2024 (.NET Framework 4.8)">
  <RuntimeRequirements OS="Win64" Platform="PLNT3D" SeriesMin="R24.3" SeriesMax="R24.3" />
  <ComponentEntry AppName="WeldPropUtils" ModuleName="./Contents/net48/WeldPropUtils.dll" LoadOnCommandInvocation="True">
    <Commands GroupName="WELDPROP"><Command Local="SetWeldProp" Global="SetWeldProp" /><Command Local="SetWeldNumber" Global="SetWeldNumber" /></Commands>
  </ComponentEntry>
</Components>
<Components Description="Plant 3D 2025-2026 (.NET 8)">
  <RuntimeRequirements OS="Win64" Platform="PLNT3D" SeriesMin="R25.0" SeriesMax="R25.1" />
  <ComponentEntry AppName="WeldPropUtils" ModuleName="./Contents/net8/WeldPropUtils.dll" LoadOnCommandInvocation="True"> … </ComponentEntry>
</Components>
```
Install the bundle to `%ProgramData%\Autodesk\ApplicationPlugins\` (all users) or `%AppData%\Autodesk\ApplicationPlugins\`.
The `SeriesMin/Max` values and `Platform="PLNT3D"` have to be confirmed on a real install. Treat them as the starting point.

## Checklist when adding a Plant 3D version
1. Get its SDK onto the Drive. Add its file IDs to `plant3d-build-check`, and a target plus a `PLANT_REF_<year>` variable to `tools/compile-check`.
2. Diff its API surface against the previous version (see `plant3d-api` → *Looking up other members*). Record the differences in `plant3d-api`.
3. Add the TFM to the real csproj and a `<Components>` block to the bundle.
4. Ask the user to test NETLOAD and both commands in that Plant 3D version.
