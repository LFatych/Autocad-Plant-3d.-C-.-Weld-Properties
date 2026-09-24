# WeldPropUtils — AutoCAD Plant 3D plugin

A C# class library that runs inside **AutoCAD Plant 3D**. It reads the parts on each side of a weld (a Plant 3D
*connector*) and copies their properties into custom properties on the weld. It can also assign weld numbers.

Target versions: **2024** (.NET Framework 4.8), **2026** (.NET 8), and later **2027+** (.NET 10).
The csproj is SDK-style and multi-targeted: `net8.0-windows` (2025/2026) and `net48` (2024). Each target builds only when its SDK
variable is set (see *Build and run*). The user runs **Plant 3D 2026 only**, so net8 is the main target and the one they test.

## Layout

```
WeldPropUtils/WeldPropUtils.sln
WeldPropUtils/WeldPropUtils/
  WeldPropertiesHandler.cs  commands (SetWeldProp, SetWeldNumber) + the loop over connectors
  MiscUtilities.cs          Acad (static doc/db/editor/DataLinksManager) + extension helpers
  WeldKey.cs                structPort + Weld model (port normalisation); structPort.Props = all part properties
  Settings/                 per-project JSON settings (WeldPropSettings model + SettingsStore load/save)
  WeldPropUtils.csproj      SDK-style; net8.0-windows from $(AP3D_SDK_2026), net48 from $(AP3D_SDK_2024)
  Properties/launchSettings.json   F5 starts Plant 3D 2026 (acad.exe /product PLNT3D)
tools/compile-check/        Linux compile check against the real 2024 + 2026 SDK DLLs (used by Claude)
.claude/skills/             domain knowledge for Claude (see below)
```

## Commands

| Command         | What it does |
|-----------------|--------------|
| `SetWeldProp`   | For every visible `ACPPCONNECTOR` whose `JointType` is `Buttweld`, `Tap` or `Socketweld`, writes `Material1/2`, `OD1/2`, `WallThickness1/2`, `LDS1/2` and `SPEC1/2` onto the weld sub-part row. |
| `SetWeldNumber` | Runs `SetWeldProp`, then groups welds that have the same OD, wall thickness and material on both ports. It numbers each group, starting from 11 for butt welds, 51 for taps and 71 for socket welds. Writes the result to `WeldNumber`. |

Both commands read `<Plant project folder>\SPDS\WeldPropUtils.json` (see *Project settings* below).

Port 1 is always the "larger" side (`Weld.NormalizePorts`: OD, then wall thickness, then material).

### Project prerequisites (in the Plant 3D project, not in code)
The weld class must have these custom properties, added in Project Setup:
`Material1, OD1, WallThickness1, LDS1, SPEC1, Material2, OD2, WallThickness2, LDS2, SPEC2, WeldNumber`.

### Project settings (`<project>\SPDS\WeldPropUtils.json`)
- Location: `PlantProject.ProjectFolderPath` (the folder containing `Project.xml`) + `SPDS\WeldPropUtils.json`.
  Each Plant project has its own file. Other SPDS tools should add their own files to the same `SPDS` folder.
- Created with the defaults (the same behaviour as the original hard-coded plugin) the first time a command runs.
  A file that can't be read is left untouched and the defaults are used for that run, with a message on the command line.
- Content: `schemaVersion`, `activeProfile`, `autoUpdate` (reserved for step 2), `numbering` (start numbers for
  Buttweld/Tap/Socketweld), and `profiles[]`, each with `name`, `mirrorSides` and `mappings[]` of `{target, side (1|2), source}`.
- Serializer: `DataContractJsonSerializer`, which is built into net48 and net8. Don't add Newtonsoft.Json, because AutoCAD loads its own copy.
  When adding a member: `[DataMember(Name = "camelCase")]`, a default in `Normalize()`, and raise `schemaVersion` if old
  files need converting.
- Per-user preferences, such as the UI theme, don't belong here. They go to `%AppData%\SPDS\WeldPropUtils\`.

## Build and run (Windows)
- Visual Studio 2022 (17.8+) with the .NET 8 SDK. The .NET Framework 4.8 targeting pack is only needed for the 2024 build.
- Environment variables, set to the SDK root folders (restart Visual Studio after setting them):
  - `AP3D_SDK_2026` → the Plant 3D 2026 SDK folder (Drive `AP3D_2026_SDK`). This builds `net8.0-windows`.
  - `AP3D_SDK_2024` → the Plant 3D 2024 SDK folder. This is optional and builds `net48`.
  The csproj takes `inc\AcCoreMgd.dll`, `AcDbMgd.dll` and `AcMgd.dll` and every `inc-x64\PnP*.dll` from there, with `Private=False`.
- Output: `bin\<Config>\net8.0-windows\WeldPropUtils.dll` (2025/2026) and `bin\<Config>\net48\WeldPropUtils.dll` (2024).
- F5 starts Plant 3D 2026 (`Properties/launchSettings.json`). Load the DLL with `NETLOAD`. A Plant 3D project must be open.

## Working with Claude in this repo
- **Claude can compile but not run.** `tools/compile-check` builds the sources on Linux against the real SDK
  DLLs for `net48` (2024) and `net8.0-windows` (2026), by building the real csproj. The DLLs come from the user's Google Drive and are **never
  committed** (see the `plant3d-build-check` skill). The user then tests behaviour in Plant 3D. Every change
  ends with a short manual test plan.
- Only use API members listed in `.claude/skills/plant3d-api/reference.md`, or ones you verified in the SDK.
  Never invent members.
- The shared code must compile as **C# 7.3** (the net48 target) and on .NET 8. No .NET Framework-only APIs.
- Match the existing style: extension methods in `MiscUtilities`, commands in `WeldPropertiesHandler`.
- Skills in `.claude/skills/`:
  - `plant3d-api`: verified Plant 3D API facts + `reference.md` with exact signatures
  - `plant3d-build-check`: fetch the SDK DLLs from Drive and compile-check both targets
  - `plant3d-multi-version`: 2024/2026/2027 targeting, csproj layout, autoloader bundle
  - `autocad-net-plugin`: AutoCAD .NET basics (commands, transactions, locking, selection, loading)
  - `plant3d-change-review`: the checklist before every commit
  - `spds-wpf-ui`: SPDS brand themes (dark + light), fonts, logo, WPF hosting in AutoCAD, approved mapping-window layout
- Roadmap (agreed with the user):
  1. ✅ Per-project settings file with mapping profiles (the commands already use it).
  2. WPF **Weld Property Mapping** window that edits those profiles (design: `spds-wpf-ui`), plus the manual "update all welds".
  3. Auto-update through events: `DataLinksManager.DataLinkOperationOccurred` only *collects* the affected row IDs
     (no DB work inside the handler), then `Document.CommandEnded` processes them once. Include a re-entrancy guard,
     skip UNDO/REDO/sync/audit, and an on/off switch (`autoUpdate`). Weld numbering stays manual.
- The SessionStart hook (`.claude/hooks/session-start.sh`) installs `dotnet-sdk-8.0` in cloud sessions.

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
