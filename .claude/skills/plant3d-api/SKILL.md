---
name: plant3d-api
description: Verified reference for the AutoCAD Plant 3D .NET API (Autodesk.ProcessPower.*) for Plant 3D 2024 (.NET Framework 4.8), 2026 (.NET 8) and 2027 (.NET 10) — DataLinksManager, row IDs vs ObjectIds, PpObjectId/sub-indexes, PnPDatabase/PnPRow, Connector/WeldSubPart, Equipment/NozzleSubPart, Port/Pair/ConnectionManager, property names, and the project database tables. Use whenever reading or writing Plant 3D object data, walking piping connectivity, handling welds or nozzles, or adding a command that touches Plant 3D parts.
---

# AutoCAD Plant 3D .NET API

Checked against the user's SDKs (Google Drive folder `1S-v3uK1nUuRrxSziWlyyNFYntvLAh918`):
- `AP3D_2024_SDK`: assemblies version 15.0, .NET Framework 4.7/4.8.
- `AP3D_2026_SDK`: assemblies version 17.0, `.NETCoreApp,Version=v8.0`.
- `AP3D_2027_SDK` (folder `1bHCL2xQfAR9JbWepFa0TAloD7PHH1boo`): assemblies version 18.0 (AcCoreMgd 26.0), `.NETCoreApp,Version=v10.0`.
  **2027 is the default build target of this repo.**

Sources: the developer guide `docs/plantsdk_dev.chm` (read in full), the reference `docs/plantsdk_ref.chm`, and the
metadata of the `inc-x64/*.dll` files.
**`reference.md` next to this file lists the exact member signatures.** Use it before writing any call.
Plant 3D 2028 and later: no SDK yet. Treat anything there as unverified.

## Assemblies → namespaces

| DLL | Namespace | Key types |
|---|---|---|
| `inc\AcCoreMgd.dll`, `AcDbMgd.dll`, `AcMgd.dll` | `Autodesk.AutoCAD.*` | Document, Database, Editor, Transaction |
| `inc-x64\PnPProjectManagerMgd.dll` | `Autodesk.ProcessPower.PlantInstance`, `.ProjectManager` | `PlantApplication`, `PlantProject`, `Project`, `PnPProjectUtils` |
| `inc-x64\PnP3dProjectPartsMgd.dll` | `Autodesk.ProcessPower.P3dProjectParts` | `PipingProject` |
| `inc-x64\PnPDataLinks.dll` | `Autodesk.ProcessPower.DataLinks` | `DataLinksManager`, `PpObjectId`, `PpObjectIdArray`, `DLException`, `DLStatus` |
| `inc-x64\PnPDataObjects.dll` | `Autodesk.ProcessPower.DataObjects` | `PnPDatabase`, `PnPTable`, `PnPRow`, `PnPRowIdArray` |
| `inc-x64\PnP3dObjectsMgd.dll` | `Autodesk.ProcessPower.PnP3dObjects` | `Part`, `Pipe`, `Connector`, `Equipment`, `SubPart`, `WeldSubPart`, `NozzleSubPart`, `Port`, `Pair`, `ConnectionManager`, `ConnectionIterator` |
| `inc-x64\PnPCommonMgd.dll` | `Autodesk.ProcessPower.Common` | shared helpers (needed by the others at compile time) |
| `inc-x64\PnP3dStructureObjectsMgd.dll` (**2027+**) | `Autodesk.ProcessPower.PnP3dStructureObjects` | `Structure`, `StructureElement`, `StructureMember`, `StructurePlate`, `StructureGrating`, `StructureFooting`, `Catalog`, `Material`, `ShapeSize`, `ShapeStandard`, `ShapeType`, `Justification`, `MiterType` (not referenced by this repo yet) |

Watch out for name clashes. `Part`, `Port` and `OpenMode` exist in more than one namespace. This repo uses the aliases
`pPart`, `pPort` and `portCol`.

## Project and DataLinksManager

```csharp
PlantProject plant = PlantApplication.CurrentProject;                 // null if no project open
Project piping = plant?.ProjectParts["Piping"];                       // PipingProject; check piping.Isloaded()
DataLinksManager dlm = piping.DataLinksManager;
PnPDatabase db = dlm.GetPnPDatabase();
// or, for "the project part that owns the active drawing":
Project prj = PnPProjectUtils.GetProjectPartForCurrentDocument();     // "Piping", "PnId", "Ortho", "Iso"
string type = PnPProjectUtils.GetActiveDocumentType();
```
Resolve these **inside each command**, never in static field initialisers. The user can switch drawings or projects.
`DataLinksManager.GetManager(Database)` returns the drawing's own data blob, which is read-only until the drawing is saved.
Use the project's DLM for anything you change.

## Identifiers

- **ObjectId**: unique within one DWG only.
- **PpObjectId**: a struct of `DwgId`, `dbHandle` and `SubIndex`. It is unique across the whole project.
  `dlm.MakeAcPpObjectId(id)` is the same as `MakeAcPpObjectId(id, 0)`. **Sub-index 0 means the object itself.**
  Sub-entities (the weld inside a connector, the nozzles on equipment) use sub-index ≥ 1. The docs say "returns nozzle
  subentity id on success, zero on failure".
- **Row ID** (`int`, shown as PnPID in Data Manager): the key of the object's row in the project database.
  `dlm.FindAcPpRowId(ObjectId | PpObjectId)`.
  **It throws `DLException` when there is no link. It does not return 0.** Check `dlm.HasLinks(oid)` first or catch the exception.
  One row can link to many objects (e.g. across drawings). Each object links to at most one row.
- List every sub-id of an object: `dlm.SelectObjectSubIds(dlm.MakeAcPpObjectId(objId))` → `PpObjectIdArray`,
  each with its `SubIndex`. This is the reliable way to find sub-part rows without guessing indexes.

## Reading and writing properties

```csharp
List<KeyValuePair<string,string>> all = dlm.GetAllProperties(rowId, true);   // bool = bCurrentVersion
StringCollection vals = dlm.GetProperties(rowId, names, true);               // bool = bCurrentVersion
dlm.SetProperties(rowId, names, values);                                      // preferred write path
bool has = dlm.HasProperty(rowId, "WeldNumber");
string cls = dlm.GetObjectClassname(rowId);                                   // e.g. "Buttweld", "Pipe"
// Direct row access (sample-style; bypasses DataLinks bookkeeping — prefer SetProperties):
PnPRow row = db.GetRow(rowId); row.BeginEdit(); row["WeldNumber"] = "11"; row.EndEdit();
```
- All three overloads accept `ObjectId`, `PpObjectId` or `int rowId`.
- Names are **internal column names** (`MatchingPipeOd`), not the display names in the palette.
- Values are strings. `GetAllProperties` gives back every column, including empty ones. This repo drops empty values.
- A custom property must be added in Project Setup on the right class (or a parent class) before you can write it.
  Otherwise `SetProperties` throws `DLException` (`FailedToSetProperties`).

### Property names used in this repo
| Where | Name | Meaning |
|---|---|---|
| Connector row | `JointType` | `Buttweld`, `Tap`, `Socketweld`, … |
| Part row | `Material`, `Spec`, `WallThickness`, `MatchingPipeOd`, `PartSizeLongDesc` | copied to the weld |
| Nozzle row | `PortName` | nozzle's port name on the equipment |
| Weld sub-part row | `Material1/2`, `OD1/2`, `WallThickness1/2`, `LDS1/2`, `SPEC1/2`, `WeldNumber` | **custom** project properties |

### Piping database tables (class names)
These come from the developer guide's list for a piping project. Weld-related classes: `Buttweld`, `Socketweld`, `FusionWeld`, `TapWeld`,
`P3dConnector`, `Gasket`, `BoltSet`. Parts: `Pipe`, `Elbow`, `Tee`, `Reducer`, `Flange`, `Valve`, `Nozzle`, `Olet`, …
Equipment: `Equipment`, `Pump`, `Vessel`, `Tank`, … Line numbers: `P3dLineGroup` (column `Number`).
Relationships worth knowing:
- `P3dLineGroupPartRelationship` (roles `LineGroup`, `Part`)
- `P3dPartConnection` (roles `Part1`, `Part2`, with `Port1`/`Port2` values)
- `PartPort`
- `AssetOwnership` (roles `Owner`, `Owned`; this is how equipment owns its nozzles)

Query one with `dlm.GetRelatedRowIds(relType, role1, rowId, role2)`.
Listing a table: `db.Tables["Pipe"].Select("filter")`. Note that `PnPTable.Rows` holds only rows pending commit.

## 3D entities
DXF names for `SelectionFilter`:
- `ACPPCONNECTOR`: connectors, meaning welds, gasket and bolt sets, etc. (managed class `Connector`)
- `ACPPPIPE`
- `ACPPPIPEINLINEASSET`
- `ACPPEQUIPMENT`

Class tree: `Part : Entity` → `Connector`, `Segment` → `Pipe`, `InlineAsset` → `PipeInlineAsset`, `Equipment`, `Support`.
The sub-parts of a `Connector` are `SubPart`s. A weld's is `WeldSubPart : JointMarkerSubPart` (it has a `Width`); bolt sets use `BoltSetSubPart`.
`Connector.NumSubParts`, `GetSubPart(i)` and `AllSubParts` are indexed from 0 on the entity. The DataLinks sub-index is ≥ 1 (see Identifiers).
`Equipment.AllSubParts` holds `NozzleSubPart` items, and each has an `Index` property.
`Part.PartSizeProperties.PropValue("…")` reads spec/catalog values straight from the entity, without the database.

## Connectivity

```csharp
PortCollection ports = part.GetPorts(PortType.Both);        // Static=1, Dynamic=2, Both=3, Symbolic=4, All=7
var cm = new ConnectionManager();
var me = new Pair { ObjectId = part.ObjectId, Port = port };
if (cm.IsConnected(me)) {
    Pair other = cm.GetConnectedPairAt(me);                  // other part's ObjectId + its Port (+ PpObjectId incl. SubIndex)
}
ConnectionCollection all = cm.GetConnections(part);         // Connection.Pair1/Pair2, Port1/Port2
ConnectionIterator it = ConnectionIterator.NewIterator(part.ObjectId, port);
for (; !it.Done(); it.Next()) { ObjectId id = it.ObjectId; }
```
- The SDK sample `PipingValidation` (UnconnectedPortRule) uses `Pair` together with `IsConnected`.
- `GetConnectedPairAt` gives you the other side's port straight away. You don't need to compare `Port.Position` values.
  If you do compare positions, use `Point3d.IsEqualTo` (with a tolerance), not `==`.
- Equipment: the connected `Pair.Port.Name` is an equipment port name (the guide shows `S10002`, where plain parts use `S1`/`S2`).
  To reach the nozzle row, try `cm.GetConnectedPairAt(me).PpObjectId`. When its `SubIndex > 0`,
  `dlm.FindAcPpRowId(thatPpId)` should be the nozzle row. **This still needs a runtime test.** The fallback is
  `SelectObjectSubIds` on the equipment plus matching each sub-row's `PortName`.

## Error handling
- `DLException` (namespace DataLinks) carries a `DLStatus`, for example `NotFound=6`, `ObjectDoesNotHaveLink=16`,
  `FailedToSetProperties=22` or `FailedToGetProperties=23`.
- `PnP3dException` and its subclasses (`InvalidConnectorException`, …) come from PnP3dObjects.

## 2024 → 2026 differences (API surface diff of the six DLLs)
Everything this repo uses is **identical**. The differences are:
- `PnPDataObjects`: rebuilt from C++/CLI into managed code. Base types were reorganised, but `PnPRow` still has the
  `this[string]` indexer (inherited from `PnPAbstractRow`) and `BeginEdit/EndEdit`. `PnPCounter` is no longer `IDisposable`.
- `PipingProject.Audit()` became `Audit(bool bAutoCheckin, bool bShowProgress)`.
- `PnPCommonMgd`: Excel export additions. `DataLinks.ProductInformation` is new (`ProductYear`, `ProductVersion`).
- `IntPtr` became `nint` (not visible to callers).

Proof: `tools/compile-check` builds the unchanged repo code for both SDKs with no errors (see the `plant3d-build-check` skill).

## 2026 → 2027 differences (API surface diff of the DLLs)
Everything this repo uses is **identical** (PnP3dObjectsMgd and PnP3dProjectPartsMgd have no changes at all). The differences are:
- `DataLinksManager.BeginMultiDrawingMerge()` / `EndMultiDrawingMerge()` (new; "flag files for multi-drawing merge").
  Not documented in detail yet; verify before use.
- New `PnP3dStructureObjectsMgd.dll`: the structure API (members, plates, gratings, footings, catalog/shape data).
- `PnPDataObjects`: `PnPDatabaseLink.ProviderPasswordCallback`, `PnPQryParser.ColumnNeedsQuote`; `PnPSortItem` became a
  primary-constructor struct (same constructor).
- `PnPProjectManagerMgd`: document-management helpers lost their `recursive` parameters
  (`DocumentManagementUtils.CollectFileAssociations/CollectFilesForCheckOut`), `ValidateLocalWorkspace(bool bCheckPermissions)`,
  new `PickListConflict`, `ObjectConflict.Clone()`. `AcquisitionStatus.NoOrMultipleSource` removed.
- `PnPCommonMgd`: small Excel/import additions; `ShareType.IPC` removed.
- `ProductInformation.ProductYear` = "2027", `ProductVersion` = "18.0.x".
- Other 2027 "What's New" items (developer guide): active tool palette, Navisworks/Point Cloud dictionaries, xref block
  table record handle, highlight/unhighlight, collect xrefs.
- Runtime: .NET 10. Building needs the .NET 10 SDK, and Autodesk's guide asks for **Visual Studio 2026 (18.x)**.

## Looking up other members
The decompiled SDK is not in the repo. To look up another member:
1. Fetch the DLL (see `plant3d-build-check` for the Drive file IDs).
2. Run `ilspycmd -o <dir> <dll>`: `dotnet tool install -g ilspycmd` (11.x; 8.x fails on the net10 DLLs), then run it with
   `DOTNET_ROLL_FORWARD=Major`.
3. Grep the decompiled output.

For the reference help, run `7z x plantsdk_ref.chm` and grep the HTML files.
Never add Autodesk DLLs or decompiled sources to git.
