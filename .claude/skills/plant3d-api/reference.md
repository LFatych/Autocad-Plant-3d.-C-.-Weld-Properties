# Plant 3D .NET API — verified signatures

Extracted from the public metadata of the Plant 3D **2024** SDK DLLs (`inc-x64`, assembly version 15.0).
The same members exist unchanged in the **2026** SDK (17.0, .NET 8) and the **2027** SDK (18.0, .NET 10), except where noted in
SKILL.md ("2024 → 2026 differences", "2026 → 2027 differences").
Modifiers such as `unsafe`/`virtual` and marshalling attributes are stripped. `bool` parameters are plain `bool` in C#.
This is a curated subset. To look up anything else, see *Looking up other members* in SKILL.md.

## DataLinks.DataLinksManager (PnPDataLinks.dll)

```
<class> public class DataLinksManager : IDisposable
List<KeyValuePair<string, string>> GetAllProperties(ObjectId oid, bool bCurrentVersion)
List<KeyValuePair<string, string>> GetAllProperties(PpObjectId oid, bool bCurrentVersion)
List<KeyValuePair<string, string>> GetAllProperties(int rowid, bool bCurrentVersion)
ObjectId MakeAcDbObjectId(PpObjectId id)
ObjectId MakeAcDbObjectId(PpObjectId id, Database db)
ObjectIdCollection MakeAcDbObjectIds(PpObjectId id)
PnPDatabase GetPnPDatabase()
PnPRowIdArray GetDrawingIdsOutOfSync()
PnPRowIdArray GetRelatedRowIds(int rid)
PnPRowIdArray GetRelatedRowIds(string relationshipType, string role1, int rid, string role2)
PnPRowIdArray SelectAcPpRowIds(Database db)
PnPRowIdArray SelectAcPpRowIds(Database db, string classname)
PnPRowIdArray SelectAcPpRowIds(PpObjectIdArray oids)
PnPRowIdArray SelectAcPpRowIds(int dwgid)
PnPRowIdArray SelectAcPpRowIds(int dwgid, string classname)
PpObjectId MakeAcPpObjectId(ObjectId id)
PpObjectId MakeAcPpObjectId(ObjectId id, int subindex)
PpObjectIdArray FindAcPpObjectIds(int dbid)
PpObjectIdArray FindAcPpObjectIds(int dbid, bool bInMemoryOnly)
PpObjectIdArray GetRelatedAcPpObjectIds(PpObjectId oid)
PpObjectIdArray GetRelatedAcPpObjectIds(string relationshipType, string role1, PpObjectId oid, string role2)
PpObjectIdArray SelectObjectSubIds(PpObjectId oid)
StringCollection GetProperties(ObjectId oid, StringCollection pnames, bool bCurrentVersion)
StringCollection GetProperties(PpObjectId oid, StringCollection pnames, bool bCurrentVersion)
StringCollection GetProperties(int rowid, StringCollection pnames, bool bCurrentVersion)
bool HasLinks(ObjectId oid)
bool HasLinks(PpObjectId oid)
bool HasLinks(int rowid)
bool HasLinks(int rowid, bool bInMemoryOnly)
bool HasLinksInDwg(int rowid, int dwgid)
bool HasProperty(ObjectId oid, string iPropertyName)
bool HasProperty(PpObjectId oid, string iPropertyName)
bool HasProperty(int rowId, string iPropertyName)
int CreateLinkSetProperties(ObjectId oid, string classname, StringCollection pnames, StringCollection pvals)
int CreateLinkSetProperties(PpObjectId oid, string classname, StringCollection pnames, StringCollection pvals)
int FindAcPpRowId(ObjectId oid)
int FindAcPpRowId(PpObjectId oid)
int GetDrawingId(Database db)
int GetDrawingId(string fguid)
static DataLinksManager GetManager(Database A_0)
static DataLinksManager GetManager(string name)
static DataLinksManager GetManager(string name, PnPDatabaseLink dblink, bool bUsePersistentCache, bool bForceDropPersistentCache, bool bBuildPersistentCache)
static DataLinksManager GetManager(string name, Stream stream)
static DataLinksManager GetManager(string name, string username, string password)
string GetObjectClassname(ObjectId oid)
string GetObjectClassname(PpObjectId oid)
string GetObjectClassname(int oid)
void Relate(ObjectId oid1, ObjectId oid2)
void Relate(PpObjectId oid1, PpObjectId oid2)
void Relate(int rid1, int rid)
void Relate(string relationshipType, string role1, ObjectId oid1, string role2, ObjectId oid2)
void Relate(string relationshipType, string role1, PpObjectId oid1, string role2, PpObjectId oid2)
void Relate(string relationshipType, string role1, int rowid1, string role2, int rowid2)
void SetProperties(ObjectId oid, StringCollection pnames, StringCollection pvals)
void SetProperties(PpObjectId oid, StringCollection pnames, StringCollection pvals)
void SetProperties(int rowid, StringCollection pnames, StringCollection pvals)
void Unrelate(ObjectId oid1, ObjectId oid2)
void Unrelate(PpObjectId oid1, PpObjectId oid2)
void Unrelate(int rid1, int rid2)
void Unrelate(string relationshipType, string role1, ObjectId oid1, string role2, ObjectId oid2)
void Unrelate(string relationshipType, string role1, PpObjectId oid1, string role2, PpObjectId oid2)
void Unrelate(string relationshipType, string role1, int rowid1, string role2, int rowid2)
```

## DataLinks.PpObjectId (struct)

```
<struct> public struct PpObjectId : IComparable<PpObjectId>
PpObjectId(int dbid, long handle)
PpObjectId(int dbid, long handle, int sindex)
bool IsNull
int CompareTo(PpObjectId other)
int DwgId
int SubIndex
long dbHandle
static PpObjectId Null;
```

## DataLinks.DLException / DLStatus

```
FindAcPpRowId, GetProperties, SetProperties etc. throw DLException when the native call
returns a status other than DLStatus.Ok (e.g. NotFound = 6, ObjectDoesNotHaveLink = 16,
FailedToSetProperties = 22, FailedToGetProperties = 23). They do NOT return 0/-1 for 'not linked'.
```

## PnP3dObjects.Part (PnP3dObjectsMgd.dll)

```
<class> public abstract class Part : Entity
Part()
PartSizeProperties PartSizeProperties
Point3d CenterOfGravity
Point3d Position
Port FindPort(string name, PortType type)
PortCollection GetPorts(PortType type)
SpecPort PortProperties(string name)
bool AddPort(Port port)
bool IsAnchored
bool IsUserCenterOfGravity
bool PlacementLock
bool RemovePort(string name)
bool SetEngagementLength(string portName, double length)
string GenerateDynamicPortName(bool bSymbolic)
```

## PnP3dObjects.Port (PnP3dObjectsMgd.dll)

```
<class> public class Port : DisposableWrapper
NominalDiameter NominalDiameter
Point3d Position
Port()
Vector3d Direction
bool IsDynamic
bool IsNull
bool IsSymbolic
double EngagementLength
string EndType
string Name
void SetNull()
```

## PnP3dObjects.PortCollection (PnP3dObjectsMgd.dll)

```
<class> public class PortCollection : PnP3dCollection
Port this[int index]
Port this[string name]
```

## PnP3dObjects.PortType (PnP3dObjectsMgd.dll)

```
<enum> public enum PortType
enum All = 7
enum Both = 3,
enum Dynamic = 2,
enum Static = 1,
enum Symbolic = 4,
```

## PnP3dObjects.Connector (PnP3dObjectsMgd.dll)

```
<class> public class Connector : Part
Connector()
Matrix3d Ecs
SubPart GetSubPart(int iIndex)
SubPartCollection AllSubParts
Vector3d XAxis
Vector3d ZAxis
bool GetOverrideLastPort()
bool SetOrientation(Vector3d xAxis, Vector3d zAxis)
double InsulationDiameter
double OffsetTolerance
double SlopeTolerance
int NumSubParts
void AddSubPart(SubPart subpart)
void ClearOverrideLastPort()
void RemoveSubPart(int iIndex)
void SetOverrideLastPort(Point3d position, Vector3d direction)
```

## PnP3dObjects.SubPart (PnP3dObjectsMgd.dll)

```
<class> public abstract class SubPart : Drawable
ObjectId Id
Part Parent
PartSizeProperties PartSizeProperties
Port FindPort(string name, PortType type)
PortCollection GetPorts(PortType type)
bool IsPersistent
bool MirrorFlag
double Rotation
```

## PnP3dObjects.SubPartCollection (PnP3dObjectsMgd.dll)

```
<class> public class SubPartCollection : PnP3dCollection
SubPart this[int iIndex]
```

## PnP3dObjects.JointMarkerSubPart (PnP3dObjectsMgd.dll)

```
<class> public class JointMarkerSubPart : SubPart
JointMarkerSubPart()
```

## PnP3dObjects.WeldSubPart (PnP3dObjectsMgd.dll)

```
<class> public class WeldSubPart : JointMarkerSubPart
WeldSubPart()
double Width
```

## PnP3dObjects.BlockSubPart (PnP3dObjectsMgd.dll)

```
<class> public class BlockSubPart : SubPart
BlockSubPart()
ObjectId SymbolId
```

## PnP3dObjects.NozzleSubPart (PnP3dObjectsMgd.dll)

```
<class> public class NozzleSubPart : BlockSubPart
Matrix3d Ecs
NozzleSubPart()
int Index
```

## PnP3dObjects.NozzleSubPartCollection (PnP3dObjectsMgd.dll)

```
<class> public class NozzleSubPartCollection : PnP3dCollection
NozzleSubPart this[int iIndex]
```

## PnP3dObjects.Equipment (PnP3dObjectsMgd.dll)

```
<class> public class Equipment : Part
Equipment()
Matrix3d Ecs
NozzleSubPart GetSubPart(int index)
NozzleSubPartCollection AllSubParts
ObjectId SymbolId
Vector3d XAxis
Vector3d ZAxis
bool AddSubPart(NozzleSubPart newSubPart)
bool AddSubPart(int subIndex, NozzleSubPart newSubPart)
bool RemoveAllSubParts()
bool RemoveSubPart(Port port)
bool RemoveSubPart(int subIndex)
double Rotation
int NumSubParts
```

## PnP3dObjects.Pipe (PnP3dObjectsMgd.dll)

```
<class> public class Pipe : Segment
Matrix3d Ecs
ObjectId Decoration
Pipe()
Point3d CenterOfGravity
Point3d EndPoint
Point3d StartPoint
bool FixedLength
double InsulationDiameter
double Length
double OuterDiameter
```

## PnP3dObjects.InlineAsset (PnP3dObjectsMgd.dll)

```
<class> public abstract class InlineAsset : Part
InlineAsset()
Matrix3d Ecs
ObjectId SymbolId
Vector3d XAxis
Vector3d ZAxis
bool SetOrientation(Vector3d xAxis, Vector3d zAxis)
double InsulationDiameter
```

## PnP3dObjects.PipeInlineAsset (PnP3dObjectsMgd.dll)

```
<class> public class PipeInlineAsset : InlineAsset
PipeInlineAsset()
```

## PnP3dObjects.Support (PnP3dObjectsMgd.dll)

```
<class> public class Support : Part
Matrix3d Ecs
ObjectId SymbolId
Support()
Vector3d XAxis
Vector3d ZAxis
double Rotation
```

## PnP3dObjects.PartSizeProperties (PnP3dObjectsMgd.dll)

```
<class> public class PartSizeProperties : DisposableWrapper
NominalDiameter NominalDiameter
PartSizeProperties()
PartSizeProperties(PartSizeProperties src, bool bDetach)
SpecPort Port(string portName)
SpecPort PrincipalPort
StringCollection PortNames
StringCollection PropNames
Units LinearUnit
bool RenamePort(string oldName, string newName)
double NeedUnitScale(Units targetUnit)
int PartId
int PortCount
int PropCount
object PropValue(string name)
static object FormatSizeDisplay(string size, Units targetUnit, bool bIncludeNativeString, Units partUnit)
static object SizeFromDisplay(string displaySize)
static string LengthUnitPropertyName
string Definition
string Domain
string Name
string OriginalVersion
string PartGuid
string Spec
string Type
void SetPropValue(string name, object value)
```

## PnP3dObjects.Pair (PnP3dObjectsMgd.dll)

```
<class> public class Pair : DisposableWrapper
ObjectId ObjectId
Pair()
Port Port
PpObjectId PpObjectId
bool IsDrawingOnly
bool IsUnresolved
void Clear()
```

## PnP3dObjects.Connection (PnP3dObjectsMgd.dll)

```
<class> public class Connection : DisposableWrapper
Connection()
Connection(Pair pair1, Pair pair2)
ObjectId ObjectId1
ObjectId ObjectId2
Pair Pair1
Pair Pair2
Port Port1
Port Port2
bool IsDrawingOnly
bool IsExternal
void Clear()
```

## PnP3dObjects.ConnectionManager (PnP3dObjectsMgd.dll)

```
<class> public class ConnectionManager : DisposableWrapper
ConnectionCollection GetConnections(ObjectId objId)
ConnectionCollection GetConnections(Part part)
ConnectionManager()
Pair GetConnectedPairAt(Pair pair)
bool CanConnect(Pair pair1, Pair pair2)
bool IsConnected(ObjectId objId)
bool IsConnected(Pair pair)
bool IsConnected(Part part)
void Connect(Pair pair1, Pair pair2)
void Disconnect(ObjectId part1ObjId, ObjectId part2ObjId)
```

## PnP3dObjects.ConnectionIterator (PnP3dObjectsMgd.dll)

```
<class> public class ConnectionIterator : DisposableWrapper
ObjectId ObjectId
bool Done()
bool Next()
bool Next(bool bValidateConnection)
bool Start()
static ConnectionIterator NewIterator(ObjectId partId)
static ConnectionIterator NewIterator(ObjectId partId, Port port)
static ConnectionIterator NewIterator(Pair pair)
void SetStartPosition(ObjectId partId, Port port)
void SetStartPosition(Pair pair)
```

## DataObjects.PnPDatabase (PnPDataObjects.dll) — subset

```
<class> public class PnPDatabase : IDisposable
PnPRelationship[] SelectRowRelationships(int rowId)
PnPRow GetRow(int rowId)
PnPRow GetRow(int rowId, StringCollection sellist)
PnPTables Tables
static PnPDatabase Open(string connectionstring, string[] assemblies, bool bUsingPersistentCache, bool bForceUsingPersistentCache, bool bForceDropPersistentCache)
static PnPDatabase Open(string connectionstring, string[] assemblies, bool bUsingPersistentCache, bool bForceUsingPersistentCache, bool bForceDropPersistentCache, bool bCreateFullCache)
static PnPDatabase Open(string name)
static PnPDatabase Open(string name, Stream xmlstream)
static PnPDatabase Open(string name, bool bUsingPersistentCache, bool bForceUsingPersistentCache, bool bForceDropPersistentCache)
static PnPDatabase Open(string name, string userId, string password)
static PnPDatabase Open(string name, string userId, string password, bool bUsingPersistentCache, bool bForceUsingPersistentCache, bool bForceDropPersistentCache)
static PnPDatabase Open(string name, string userId, string password, bool bUsingPersistentCache, bool bForceUsingPersistentCache, bool bForceDropPersistentCache, bool bCreateFullCache)
static PnPDatabase Open(string name, string[] assemblies)
string GetRowTableName(int rowId)
string[] ExemptTables()
void CommitTransaction()
void MaterializeTables(bool bRelationships, StringCollection tables)
void RollbackTransaction()
void StartTransaction()
```

## DataObjects.PnPTable — subset

```
<class> public class PnPTable : IListSource, IDisposable
PnPRowIdArray SelectIds(string filter, PnPRowIdArray rowids)
PnPRow[] Select()
PnPRow[] Select(string filter)
string Name
```

## DataObjects.PnPRow / PnPAbstractRow

```
<class> public abstract class PnPAbstractRow
<class> public abstract class PnPRow : PnPAbstractRow, IDisposable
abstract PnPTable Table
abstract int RowId
abstract object this[PnPColumn col]
abstract object this[int columnIndex]
abstract object this[string column]
abstract string ClassName
abstract void BeginEdit();
abstract void CancelEdit();
abstract void EndEdit();
```

## PlantInstance.PlantApplication / ProjectManager.PlantProject (PnPProjectManagerMgd.dll)

```
namespace Autodesk.ProcessPower.PlantInstance
<class> public class PlantApplication
static PlantProject CurrentProject            // null when no project is open

namespace Autodesk.ProcessPower.ProjectManager   // NOT PlantInstance
<class> public class PlantProject
ProjectPartCollection ProjectParts            // indexer: Project this[string strName]  ("Piping", "PnId", "Ortho", "Iso")
string Name
string FileName                               // full path of Project.xml
string ProjectFolderPath                      // Path.GetDirectoryName(FileName): the project root folder
```

## ProjectManager.Project — subset

```
<class> public abstract class Project : XmlDocument
DataLinksManager DataLinksManager
string ProjectDirectory
bool Isloaded()                               // note the lower-case 'l'
List<PnPProjectDrawing> GetPnPDrawingFiles()
```

## ProjectManager.PnPProjectUtils — subset

```
static DwgProject DoesCurrentDwgBelongtoCurrentProject(out string sProjectPartName, out string sOtherProject)
static DwgProject DoesDocumentBelongToCurrentProject(Document oDoc, out string sProjectPartName, out string sOtherProject)
static Project GetProjectPartForCurrentDocument()
static string GetActiveDocumentType()
```

## P3dProjectParts.PipingProject (PnP3dProjectPartsMgd.dll) — subset

```
<class> public class PipingProject : Project
DataLinksManager3d DataLinksManager3d
PnPProjectSpecItem FindSpec(string name)
StringCollection GetSpecNames()
void InitializeWithDataLinksManager(DataLinksManager dlm)
void SetProjectUnitsTypeIn3dDLM(DataLinksManager3d dlm3d, ProjectUnitsType units)
```
