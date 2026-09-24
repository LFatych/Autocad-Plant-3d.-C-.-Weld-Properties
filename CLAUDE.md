# WeldPropUtils — AutoCAD Plant 3D plugin

A C# class library that runs inside **AutoCAD Plant 3D**. It reads the parts on each side of a weld (a Plant 3D
*connector*) and copies their properties into custom properties on the weld. It can also assign weld numbers.

Target: **Plant 3D 2027** (.NET 10, `net10.0-windows`), the version the user runs. One switch in the csproj,
`PlantVersion` (default `2027`), selects the target; `-p:PlantVersion=2026` still builds for Plant 3D 2026 (.NET 8).
No multi-targeting in one build (user decision: one version at a time).

## Layout

```
WeldPropUtils/WeldPropUtils.sln
WeldPropUtils/WeldPropUtils/
  WeldPropertiesHandler.cs  commands (SetWeldProp, SetWeldNumber) + the loop over connectors
  MiscUtilities.cs          Acad (static doc/db/editor/DataLinksManager) + extension helpers
  WeldKey.cs                structPort + Weld model (port normalisation); structPort.Props = all part properties
  Settings/                 per-project JSON settings (WeldPropSettings/MappingProfile, SettingsStore) + UserPreferences
  Schema/ProjectSchema.cs   reads part/weld class properties from Project Setup (project database)
  UI/                       WPF mapping window in XAML + MVVM (user decision: XAML, so it is easy to edit)
    MappingWindow.xaml(.cs)   layout + bindings; thin code-behind (theme swap, close)
    MappingViewModel.cs       window state + commands; Mvvm.cs = ObservableObject/RelayCommand
    Controls/                 custom controls: PropertyChip (drag source), DropZone (drop target)
    Themes/                   SpdsDark.xaml / SpdsLight.xaml (brushes) + SpdsStyles.xaml (styles)
  WeldPropUtils.csproj      SDK-style, PlantVersion 2027 → net10.0-windows (2026 → net8.0-windows), WPF + WinForms;
                            references from the AutoCAD <PlantVersion> install folder
  Properties/launchSettings.json   F5 starts Plant 3D 2027 (acad.exe /product PLNT3D)
tools/compile-check/        Linux compile check (C# + XAML) against the real 2027/2026 SDK DLLs (used by Claude)
.claude/skills/             domain knowledge for Claude (see below)
```

## Commands

| Command         | What it does |
|-----------------|--------------|
| `SetWeldProp`   | For every visible `ACPPCONNECTOR` whose `JointType` is `Buttweld`, `Tap` or `Socketweld`, writes the mapped weld properties of the active profile onto the weld sub-part row (default profile: `Material1/2`, `OD1/2`, `WallThickness1/2`, `LDS1/2`, `SPEC1/2`). Mapped weld properties that the weld classes don't have are skipped, with one message. |
| `WeldPropMapping` | Opens the SPDS **Weld Property Mapping** window: drag part properties (read from Project Setup) onto weld properties, add any weld class property per side ("+ Add weld property"), weld properties missing in Project Setup are shown red, manage profiles, dark/light theme. **Save** writes the settings file; **Save and update all welds** also runs the mapping on the drawing. |
| `SetWeldNumber` | Runs `SetWeldProp`, then groups welds that have the same OD, wall thickness and material on both ports. It numbers each group, starting from 11 for butt welds, 51 for taps and 71 for socket welds. Writes the result to `WeldNumber`. |

Both commands read `<Plant project folder>\SPDS\WeldPropUtils.json` (see *Project settings* below).

Port 1 is always the "larger" side (`Weld.NormalizePorts`: OD, then wall thickness, then material).

### Project prerequisites (in the Plant 3D project, not in code)
The weld classes need custom properties to write to, added in Project Setup. Any names work (e.g. the user's project has
`Port1_Material, Port1_Out_Diameter, Port1_Wall_Thickness, Port2_…`); pick them in the mapping window. The default profile
uses `Material1, OD1, WallThickness1, LDS1, SPEC1, Material2, …`. `SetWeldNumber` needs `WeldNumber`.
Side pairing (for "same mapping for both sides"): names that differ only in one `1`/`2` digit (`WeldSides.Counterpart`).
The side is stored per mapping (`side`), not derived from the name.

### Project settings (`<project>\SPDS\WeldPropUtils.json`)
- Location: `PlantProject.ProjectFolderPath` (the folder containing `Project.xml`) + `SPDS\WeldPropUtils.json`.
  Each Plant project has its own file. Other SPDS tools should add their own files to the same `SPDS` folder.
- Created with the defaults (the same behaviour as the original hard-coded plugin) the first time a command runs.
  A file that can't be read is left untouched and the defaults are used for that run, with a message on the command line.
- Content: `schemaVersion`, `activeProfile`, `autoUpdate` (reserved for step 2), `numbering` (start numbers for
  Buttweld/Tap/Socketweld), and `profiles[]`, each with `name`, `mirrorSides` and `mappings[]` of `{target, side (1|2), source}`.
- Serializer: `DataContractJsonSerializer`, which is built into .NET 8. Don't add Newtonsoft.Json, because AutoCAD loads its own copy.
  When adding a member: `[DataMember(Name = "camelCase")]`, a default in `Normalize()`, and raise `schemaVersion` if old
  files need converting.
- Per-user preferences, such as the UI theme, don't belong here. They go to `%AppData%\SPDS\WeldPropUtils\`.

## Build and run (Windows)
- **Visual Studio 2026 (18.x)** with the .NET 10 SDK (Autodesk's 2027 developer guide requires VS 2026).
  For `PlantVersion=2026`: VS 2022 17.8+ / .NET 8 SDK is enough.
- References come from the installed **AutoCAD Plant 3D 2027**: `C:\Program Files\Autodesk\AutoCAD 2027` and its `PLNT3D`
  subfolder (MSBuild `AssemblySearchPaths`), with no SDK and no environment variable. For another install folder, set `AcadDir`.
  `Private=False` (Copy Local off) on purpose: AutoCAD already loads these DLLs, and a second copy causes type-identity errors.
- Output: `bin\<Config>\net10.0-windows\WeldPropUtils.dll` (2026: `net8.0-windows`).
- F5 starts Plant 3D 2027 (`Properties/launchSettings.json`). Load the DLL with `NETLOAD`. A Plant 3D project must be open.

## Working with Claude in this repo
- **Claude can compile but not run.** `tools/compile-check` builds the sources and the XAML on Linux against the real SDK
  DLLs of Plant 3D 2027 (and 2026), by building the real csproj. The DLLs come from the user's Google Drive and are **never
  committed** (see the `plant3d-build-check` skill). The user then tests behaviour in Plant 3D. Every change
  ends with a short manual test plan.
- Only use API members listed in `.claude/skills/plant3d-api/reference.md`, or ones you verified in the SDK.
  Never invent members.
- .NET 10 / C# 14 (the SDK default); code must also compile for .NET 8 / C# 12 while 2026 is kept buildable,
  so don't use C# 13/14-only features. No .NET Framework-only APIs.
- Match the existing style: extension methods in `MiscUtilities`, commands in `WeldPropertiesHandler`.
  UI: XAML + view model + custom controls, no UI built in code (see `spds-wpf-ui`).
- Skills in `.claude/skills/`:
  - `plant3d-api`: verified Plant 3D API facts (2024/2026/2027 diffs) + `reference.md` with exact signatures
  - `plant3d-build-check`: fetch the SDK DLLs from Drive and compile-check (C# + XAML) against the 2027/2026 SDKs
  - `autocad-net-plugin`: AutoCAD .NET basics (commands, transactions, locking, selection, loading)
  - `plant3d-change-review`: the checklist before every commit
  - `spds-wpf-ui`: SPDS brand themes (dark + light), fonts, logo, WPF hosting in AutoCAD, approved mapping-window layout
- Roadmap (agreed with the user):
  1. ✅ Per-project settings file with mapping profiles (the commands already use it).
  2. ✅ (v1, untested in Plant 3D) WPF **Weld Property Mapping** window (`WeldPropMapping`). Still open: weld preview/"pick weld",
     "apply to selected welds", SPDS logo in the header (waiting for the user's OK to commit it).
  3. Auto-update through events: `DataLinksManager.DataLinkOperationOccurred` only *collects* the affected row IDs
     (no DB work inside the handler), then `Document.CommandEnded` processes them once. Include a re-entrancy guard,
     skip UNDO/REDO/sync/audit, and an on/off switch (`autoUpdate`). Weld numbering stays manual.
- The SessionStart hook (`.claude/hooks/session-start.sh`) installs `dotnet-sdk-10.0` and `dotnet-sdk-8.0` in cloud sessions.

## Known issues (from the review, verified against the SDK; not fixed yet)
1. **Nozzle lookup**: `MakeAcPpObjectId(connPart.ObjectId, 1)` always reads nozzle sub-index 1, so equipment with
   several nozzles gets the wrong nozzle's properties. `eqp` is also null when a non-equipment part has a port name longer
   than 2 characters. Fix: use `ConnectionManager.GetConnectedPairAt(pair).PpObjectId` (it carries the `SubIndex`),
   or `dlm.SelectObjectSubIds` plus a `PortName` match. Check `connPart is Equipment`, not the name length.
2. **Stale context**: `Acad` caches `doc/db/ed/dlm` in static fields the first time it is used. After switching drawings
   or projects, the commands keep acting on the old ones.
3. **Size ordering**: `Weld.ComparePorts` compares OD and wall thickness as **strings** (`"114.3" < "60.3"`).
   `SetNum` uses `Convert.ToDouble`, which depends on the locale and fails on comma-decimal Windows.
4. **Useless lock**: `using (DocumentLock ...) ;` has a stray `;` and releases immediately (compiler warning CS0642).
   Harmless for a Modal command, but misleading.
5. **Weld sub-part row**: `FindWeldRowId` always returns the row of sub-index 1, whichever sub-part is the `WeldSubPart`.
6. **Row lookups throw**: `FindAcPpRowId` throws `DLException` for unlinked objects. There's no guard, so a single
   bad object aborts the whole command.
7. Inconsistent writes: `WeldNumber` uses `PnPRow.BeginEdit/EndEdit`, other properties use `dlm.SetProperties`.
8. `GetP3dProps` shows a `MessageBox` for each failure, so a run can produce hundreds of dialogs.
9. Weld-number ranges can overlap: more than 40 butt-weld groups run into the tap range (51+).
