---
name: autocad-net-plugin
description: AutoCAD .NET (ObjectARX managed API) plugin fundamentals as they apply inside Plant 3D 2024 (.NET Framework 4.8) and 2025+ (.NET 8/10) — CommandMethod/CommandFlags, transactions, document locking, SelectionFilter/SelectAll, Editor output, IExtensionApplication, NETLOAD and debugging. Use when adding or changing commands, touching transactions or selections, or changing how the plugin is loaded.
---

# AutoCAD .NET plugin basics (inside Plant 3D)

For the target frameworks, the project file and deployment across versions, see `plant3d-multi-version`.
Reference DLLs: `AcCoreMgd` (Application, Document, Editor, CommandMethod), `AcDbMgd` (Database, Transaction, entities)
and `AcMgd` (UI). Always reference them with `Private=False`.

## Commands
```csharp
[CommandMethod("MYCMD")]                         // Modal (default): runs in document context, doc auto-locked
[CommandMethod("MYCMD", CommandFlags.UsePickSet)]// honour pre-selection (Editor.SelectImplied)
[CommandMethod("MYCMD", CommandFlags.Session)]   // application context → you MUST LockDocument()
```
- A command method can be static or an instance method. For instance methods, AutoCAD creates one object per document.
- Get the context **inside** the command:
```csharp
Document doc = Application.DocumentManager.MdiActiveDocument;
Database db = doc.Database; Editor ed = doc.Editor;
```
- Plant 3D commands need a current project. Check `PlantApplication.CurrentProject != null` and end with
  a message rather than an exception.

## Transactions
```csharp
using (Transaction tr = db.TransactionManager.StartTransaction())
{
    var part = tr.GetObject(id, OpenMode.ForRead) as Part;   // 'as' + null check for mixed selections
    // part.UpgradeOpen(); to modify the entity
    tr.Commit();   // without Commit, entity changes are rolled back
}
```
- Plant **property** writes (`DataLinksManager.SetProperties`) go to the project database, not to the DWG transaction.
  Rolling back the transaction does not undo them.
- Don't keep DBObjects after their transaction ends. Keep ObjectIds, row IDs or plain data instead.

## Document locking
You only need it in session context (modeless UI, `CommandFlags.Session`, events):
```csharp
using (DocumentLock _ = doc.LockDocument()) { /* work */ }
```
Watch for the bug `using (...) ;`. The stray `;` releases the lock immediately, and the compiler warns with CS0642.

## Selection
```csharp
var filter = new SelectionFilter(new[] {
    new TypedValue((int)DxfCode.Start, "ACPPCONNECTOR"),
});
PromptSelectionResult r = ed.SelectAll(filter);                                   // whole drawing
PromptSelectionResult r2 = ed.GetSelection(new PromptSelectionOptions(), filter); // user picks
if (r.Status != PromptStatus.OK) return;                                          // Error = nothing found
```
Use `GetSelection` or `SelectImplied` instead of `SelectAll` to let the user work on a subset.

## Messaging and UX
- Write to the command line with `ed.WriteMessage("\n...")`. Never open a `MessageBox` inside a loop.
- End with a summary, e.g. "123 welds updated, 4 skipped (unconnected)".
- Wrap each command body in try/catch (`System.Exception`, not `Autodesk.AutoCAD.Runtime.Exception` only) and
  report errors with `ed.WriteMessage`. An unhandled exception in a command can take AutoCAD down.

## Loading and debugging
- Development: run `NETLOAD` and pick the DLL. A loaded DLL cannot be unloaded, so restart Plant 3D to reload it.
- Visual Studio debugging: set the start program to `acad.exe` with the arguments `/product PLNT3D /language "en-US"` (see `.csproj.user`).
- Deployment: an autoloader `.bundle` (see `plant3d-multi-version`).
- `IExtensionApplication.Initialize()` runs when the DLL loads. Keep it light, because there may be no document or project yet.

## Culture
Number parsing has to work on any Windows locale (many use `,` as the decimal separator):
`double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var v)`.
Plant stores numeric properties as strings. Check the actual format on a real project before you rely on it.
