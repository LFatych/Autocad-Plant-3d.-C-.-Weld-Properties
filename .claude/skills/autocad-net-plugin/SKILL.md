---
name: autocad-net-plugin
description: AutoCAD .NET (ObjectARX managed) plugin fundamentals for .NET Framework 4.8 / AutoCAD 2024-based verticals — CommandMethod, CommandFlags, transactions, document locking, SelectionFilter/SelectAll, Editor messaging, IExtensionApplication, NETLOAD/autoloading, csproj references. Use when adding or changing commands, touching transactions or selections, or changing how the plugin is built or loaded.
---

# AutoCAD .NET plugin basics (AutoCAD 2024 / Plant 3D 2024)

## Target
- .NET Framework **4.8**, x64, C# 7.3 in an old-style csproj. (AutoCAD 2025+ moved to .NET 8.
  Porting to it is a separate project.)
- Reference `AcCoreMgd`, `AcDbMgd` and `AcMgd` with `Private=False`, and never copy them locally.

## Commands
```csharp
[CommandMethod("MYCMD")]                         // runs in document context; doc is auto-locked
[CommandMethod("MYCMD", CommandFlags.Modal)]     // default
[CommandMethod("MYCMD", CommandFlags.Session)]   // application context → you MUST LockDocument()
[CommandMethod("MYCMD", CommandFlags.UsePickSet)]// honour pre-selection (Editor.SelectImplied)
```
- The class holding commands may be static or instance. Instance classes get one object per document.
- Get the context **inside** the command:
```csharp
Document doc = Application.DocumentManager.MdiActiveDocument;
Database db = doc.Database; Editor ed = doc.Editor;
```

## Transactions
```csharp
using (Transaction tr = db.TransactionManager.StartTransaction())
{
    var ent = (Entity)tr.GetObject(id, OpenMode.ForRead);
    // ent.UpgradeOpen(); to write
    tr.Commit();   // without Commit everything is aborted
}
```
- Opening objects ForRead is cheap. Open for write only the objects you actually change.
- Do not keep DBObjects after the transaction ends.

## Document locking
Only needed in session context (modeless UI, `CommandFlags.Session`, events):
```csharp
using (DocumentLock _ = doc.LockDocument()) { /* work */ }
```
Watch for the bug `using (...) ;`: the lock is released straight away.

## Selection
```csharp
var filter = new SelectionFilter(new[] {
    new TypedValue((int)DxfCode.Start, "ACPPCONNECTOR"),
});
PromptSelectionResult r = ed.SelectAll(filter);            // whole drawing, no user input
PromptSelectionResult r2 = ed.GetSelection(new PromptSelectionOptions(), filter); // user picks
if (r.Status != PromptStatus.OK) return;
```
To let the user pick, or work on the pre-selection, use `GetSelection` or `SelectImplied` instead of `SelectAll`.

## Messaging & UX
- Use `ed.WriteMessage("\n...")` for command-line output. Do not open a `MessageBox` inside loops.
- Report a summary at the end, e.g. "123 welds updated, 4 skipped".
- Wrap the command body in try/catch and report errors through `ed.WriteMessage`. An unhandled exception
  can crash AutoCAD.

## Loading
- Development: use the `NETLOAD` command and pick the DLL. You cannot unload it; restart AutoCAD to reload.
- Debugging: set the Start program to `acad.exe` with `/product PLNT3D /language "en-US"`.
- Deployment: an autoloader bundle (`%AppData%\Autodesk\ApplicationPlugins\X.bundle\PackageContents.xml`)
  or registry demand-loading.
- `IExtensionApplication.Initialize()` runs when the DLL loads. Keep it light, because no document or project may be open yet.

## Culture
Windows locales differ (for example Czech, Polish, Ukrainian and German use `,` as the decimal separator). Always
`double.Parse(s, NumberStyles.Float, CultureInfo.InvariantCulture)` or `double.TryParse` with the same arguments.
