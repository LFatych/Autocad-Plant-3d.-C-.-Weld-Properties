# WeldPropUtils — AutoCAD Plant 3D plugin

A .NET Framework 4.8 class library that runs inside **AutoCAD Plant 3D 2024**. It reads the
parts on each side of a weld (a Plant 3D *connector*) and copies their properties into
custom properties on the weld. It can also assign weld numbers.

## Layout

```
WeldPropUtils/WeldPropUtils.sln
WeldPropUtils/WeldPropUtils/
  WeldPropertiesHandler.cs  commands (SetWeldProp, SetWeldNumber) + the loop over connectors
  MiscUtilities.cs          Acad (static doc/db/editor/DataLinksManager) + extension helpers
  WeldKey.cs                structPort + Weld model (port normalisation)
  WeldPropUtils.csproj      old-style csproj; references come from $(AP3D_SDK_2024)
```

## Commands

| Command         | What it does |
|-----------------|--------------|
| `SetWeldProp`   | For every visible `ACPPCONNECTOR` whose `JointType` is `Buttweld`, `Tap` or `Socketweld`, writes `Material1/2`, `OD1/2`, `WallThickness1/2`, `LDS1/2` and `SPEC1/2` onto the weld sub-part row. |
| `SetWeldNumber` | Runs `SetWeldProp`, then groups welds that have the same OD, wall thickness and material on both ports. It numbers each group, starting from 11 for butt welds, 51 for taps and 71 for socket welds. Writes the result to `WeldNumber`. |

Port 1 is always the "larger" side: `Weld.NormalizePorts` compares OD, then wall thickness, then material.

### Project prerequisites (in the Plant 3D project, not in code)
The weld class must have these custom properties, added in Project Setup:
`Material1, OD1, WallThickness1, LDS1, SPEC1, Material2, OD2, WallThickness2, LDS2, SPEC2, WeldNumber`.

## Build and run (Windows only)

- Requires Visual Studio and .NET Framework 4.8, plus the Plant 3D 2024 SDK.
- Set the environment variable `AP3D_SDK_2024` to the SDK root. The csproj takes `inc\AcCoreMgd.dll`,
  `inc\AcDbMgd.dll` and `inc\AcMgd.dll` from there, and every `PnP*.dll` from `inc-x64\`.
- All Autodesk references are `Private=False`. Never copy them to the output folder.
- Debugging launches `acad.exe /product PLNT3D` (see `.csproj.user`).
- Load the DLL with `NETLOAD`. A Plant 3D project must be open, because the static `Acad`
  class calls `PlantApplication.CurrentProject`.

## Working with Claude in this repo

- **Claude cannot compile or run this code.** The cloud container runs Linux and has neither
  the .NET Framework nor the Autodesk DLLs. Every change is checked by careful review. The user then
  builds it in Visual Studio and tests it in Plant 3D. For each change, give the user a short
  manual test plan: which drawing, which command, and what to check in the Properties palette or Data Manager.
- Only use API members that already appear in this code or in the SDK. If you are not sure a
  Plant 3D API member exists or what its exact signature is, say so. Do not invent one.
- Keep to the target: C# 7.3 (the default for net48 old-style projects). Do not use C# 8+ features
  such as `switch` expressions, `using var`, nullable reference types, or ranges.
  Note: `is null` and tuple swap work in C# 7.x. Keep the old-style csproj; do not convert to SDK-style
  unless asked.
- Match the existing style: extension methods in `MiscUtilities`, commands in `WeldPropertiesHandler`.
- The skills in `.claude/skills/` hold domain knowledge:
  - `plant3d-api`: the Plant 3D .NET API (DataLinksManager, row IDs, connectors, ports, welds, nozzles)
  - `autocad-net-plugin`: AutoCAD .NET basics (commands, transactions, locking, selection, NETLOAD)
  - `plant3d-change-review`: the checklist to run before committing any change

## Known issues (from the first review, not fixed yet)

1. Nozzle lookup always uses sub-part index `1` (`MakeAcPpObjectId(connPart.ObjectId, 1)`). Because of that,
   equipment with several nozzles can get the wrong nozzle properties. `eqp` can also be null
   when a non-equipment part has a port name longer than 2 characters.
2. `Acad` caches `doc/db/ed/dlm` in static fields when the class is first used. After switching drawings
   or projects, the commands still act on the old ones.
3. `Weld.ComparePorts` compares OD and wall thickness as **strings** (`"114.3" < "60.3"`), but
   `SetNum` sorts them as doubles using the current culture. `Convert.ToDouble` fails or gives wrong
   results on comma-decimal Windows locales.
4. `using (DocumentLock ...) ;` — the stray `;` releases the lock immediately (it does nothing).
5. `FindWeldRowId` returns the row of sub-part `1` no matter which sub-part is the `WeldSubPart`.
6. `SetNum` writes `WeldNumber` through `PnPRow.BeginEdit/EndEdit`, while other properties go through
   `dlm.SetProperties`. The approach is inconsistent.
7. `GetP3dProps` shows a `MessageBox` for each failure, which can mean hundreds of dialogs.
8. Weld-number ranges can overlap: more than 40 butt-weld groups run into the tap range (51+).
9. Build output (`bin/`, `obj/`) and `desktop.ini` were committed. `.gitignore` now excludes them.
