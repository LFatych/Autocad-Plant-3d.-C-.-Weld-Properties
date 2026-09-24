---
name: plant3d-api
description: Reference for the AutoCAD Plant 3D .NET API (Autodesk.ProcessPower.*) — DataLinksManager, row IDs, PnPDatabase, connectors/welds, ports, ConnectionManager/ConnectionIterator, equipment nozzles, sub-parts, and Plant property names. Use whenever reading or writing Plant 3D object data, walking piping connectivity, or adding a new command that touches Plant 3D parts.
---

# AutoCAD Plant 3D .NET API

> Status: written from the existing code and general knowledge of the API. It has **not** been
> checked against the user's SDK yet. Items marked ⚠ need confirmation from the SDK
> (`inc-x64/*.dll`, the SDK docs or its samples) before you rely on them. Once confirmed, update this file.

## Assemblies → namespaces (SDK 2024)

| DLL (`$(AP3D_SDK_2024)`) | Namespace(s) | Key types |
|---|---|---|
| `inc\AcCoreMgd.dll`, `AcDbMgd.dll`, `AcMgd.dll` | `Autodesk.AutoCAD.*` | Document, Database, Editor, Transaction |
| `inc-x64\PnPProjectManagerMgd.dll` | `Autodesk.ProcessPower.PlantInstance`, `.ProjectManager` | `PlantApplication`, `PlantProject`, `Project` |
| `inc-x64\PnP3dProjectPartsMgd.dll` | `Autodesk.ProcessPower.P3dProjectParts` | `PipingProject` |
| `inc-x64\PnPDataLinks.dll` | `Autodesk.ProcessPower.DataLinks` | `DataLinksManager`, `PpObjectId` |
| `inc-x64\PnPDataObjects.dll` | `Autodesk.ProcessPower.DataObjects` | `PnPDatabase`, `PnPRow`, `PnPTable` |
| `inc-x64\PnP3dObjectsMgd.dll` | `Autodesk.ProcessPower.PnP3dObjects` | `Part`, `Pipe`, `Connector`, `Equipment`, `Port`, `PortCollection`, `PortType`, `SubPart`, `WeldSubPart`, `NozzleSubPart`, `ConnectionManager`, `ConnectionIterator`, `Pair` |

Watch out for name clashes: `Part` and `Port` exist in more than one place. This code base uses these aliases:
`pPart = PnP3dObjects.Part`, `pPort = PnP3dObjects.Port`, `portCol = PnP3dObjects.PortCollection`.

## Getting the project and the DataLinksManager

```csharp
PlantProject proj = PlantApplication.CurrentProject;            // null if no project is open
PipingProject piping = proj.ProjectParts["Piping"] as PipingProject;
DataLinksManager dlm = piping.DataLinksManager;
PnPDatabase pnpDb = dlm.GetPnPDatabase();
```
Get these values **when the command runs**, not in static field initialisers. The user can switch drawings or projects.

## Two worlds: ObjectId vs. row ID
- AutoCAD entities (in the DWG) have an `ObjectId`.
- Plant data (in the project SQLite database) lives in **rows** with an `int` row ID.
- `dlm.FindAcPpRowId(ObjectId)` gives the row ID of an entity, or 0 or -1 if it is not linked ⚠.
- Sub-parts (the weld inside a connector, the nozzles on equipment) have their own rows:
  `dlm.FindAcPpRowId(dlm.MakeAcPpObjectId(ownerObjectId, subPartIndex))`.
  ⚠ Confirm in the SDK whether `subPartIndex` is 0- or 1-based and how it maps to `AllSubParts`.
  Today's code hardcodes `1`, which is suspect for equipment with several nozzles.
- Reverse lookup: `dlm.FindAcPpObjectIds(rowId)` ⚠.

## Reading and writing properties
```csharp
List<KeyValuePair<string,string>> all = dlm.GetAllProperties(rowId, true);   // 2nd arg ⚠ (meaning)
dlm.SetProperties(rowId, namesStringCollection, valuesStringCollection);       // preferred write path
// direct row access (bypasses some DataLinks behaviour — prefer SetProperties):
PnPRow row = pnpDb.GetRow(rowId); row.BeginEdit(); row["WeldNumber"] = "11"; row.EndEdit();
```
- Property **names** are internal names (`MatchingPipeOd`), not the display names shown in the palette.
- Values are strings, and numbers use an invariant `.` decimal ⚠. Parse them with `CultureInfo.InvariantCulture`.
- A custom property must first be added in Project Setup, on the right class, before code can write it.

### Property names used in this repo
| Where | Name | Meaning |
|---|---|---|
| Connector row | `JointType` | `Buttweld`, `Tap`, `Socketweld`, and also flanged/threaded values that are not welds |
| Part row | `Material`, `Spec`, `WallThickness`, `MatchingPipeOd`, `PartSizeLongDesc` | copied to the weld |
| Nozzle row | `PortName` | used to match the equipment port to the nozzle |
| Weld sub-part row | `Material1/2`, `OD1/2`, `WallThickness1/2`, `LDS1/2`, `SPEC1/2`, `WeldNumber` | **custom** project properties |

## Entities and DXF names (for SelectionFilter)
`ACPPCONNECTOR` (connector: weld, gasket and bolt set, etc.), `ACPPPIPE`, `ACPPPIPEINLINEASSET` (fittings, valves),
`ACPPEQUIPMENT`, `ACPPPIPESUPPORT` ⚠. A connector's `AllSubParts` holds a `WeldSubPart` when it is a weld.

## Connectivity
```csharp
PortCollection ports = part.GetPorts(PortType.Both);     // PortType.Static / Dynamic / Both
var cm = new ConnectionManager();
bool connected = cm.IsConnected(new Pair { ObjectId = part.ObjectId, Port = port });
ConnectionIterator it = ConnectionIterator.NewIterator(part.ObjectId, port);
for (; !it.Done(); it.Next()) { ObjectId other = it.ObjectId; /* skip Connector */ }
```
- To find the matching port on the other part, compare `Port.Position`. Use `Point3d.IsEqualTo`
  (with a tolerance), not `==`.
- Port names: plain parts use `S1`, `S2`, … Equipment ports carry the nozzle's port name, which is longer. This code
  uses `Name.Length > 2` to spot equipment. Check `part is Equipment` instead.

## Equipment nozzles
`Equipment.AllSubParts` → `NozzleSubPart` items. Each nozzle has its own row (see sub-part IDs above) whose
`PortName` matches the equipment port. Use each nozzle's own index, not a constant.

## Gotchas
- Commands need an open Plant project, or `PlantApplication.CurrentProject` is null.
- DataLinksManager writes go to the project database. Test on a copy of the project.
- In a project with a SQL Server backend, writes can fail when the user has no rights or when a row is locked.
