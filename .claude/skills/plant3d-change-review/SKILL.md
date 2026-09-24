---
name: plant3d-change-review
description: Checklist to run before committing any change to this Plant 3D plugin — compile check against the Plant 3D 2026 SDK, API/runtime/behaviour review, and the hand-off test plan for the user (who runs Plant 3D on Windows). Use before every commit, when reviewing a diff, or when the user asks "is this ready?".
---

# Pre-commit review for WeldPropUtils

Claude can compile but cannot run the plugin. Compile it, review it, then give the user a test plan.

## 1. Compiles (mandatory)
- [ ] Run the `plant3d-build-check` skill. `net8.0-windows` must build with no errors and no new warnings.
- [ ] Each new Autodesk reference is added to `WeldPropUtils.csproj` with an `$(PlantSdk)` HintPath and `Private=False`.
- [ ] Test plans are for **Plant 3D 2026** (the only supported version).

## 2. Runtime correctness (not caught by the compiler)
- [ ] Document, Database, Editor and DataLinksManager are fetched when the command runs, never cached in static fields.
- [ ] `FindAcPpRowId` is guarded, because it **throws** `DLException` when there is no link (use `HasLinks` or try/catch).
- [ ] Sub-part rows are found by the real sub-index (`SelectObjectSubIds`, `Pair.PpObjectId.SubIndex`), not a hard-coded `1`.
- [ ] `as` casts and dictionary lookups are null-checked. A missing property doesn't crash the whole run.
- [ ] Numbers are parsed with `CultureInfo.InvariantCulture`. Sizes are compared numerically, not as strings.
- [ ] No `MessageBox` inside loops. Errors are collected and reported once with `ed.WriteMessage`.
- [ ] Writes go through `dlm.SetProperties`, unless there's a reason not to (write the reason in a comment).

## 3. Behaviour
- [ ] Existing command names and outputs are unchanged unless the user asked for a change.
- [ ] Any new custom property that must exist in Project Setup is listed in the summary and in `CLAUDE.md`.

## 4. Hand-off to the user
In the final message, include:
1. What changed and why, briefly.
2. The compile-check result for each target.
3. **Build**: "Build in Visual Studio (Release), then NETLOAD `bin\Release\WeldPropUtils.dll` in Plant 3D <version>."
4. **Manual test plan**: which drawing to use and which cases to cover:
   - butt weld pipe–elbow
   - tap
   - socket weld
   - reducer (different OD on each side)
   - weld to an equipment nozzle, on equipment with **several** nozzles
   - weld with only one side connected

   Also say which command to run and which values to check in the Properties palette or Data Manager.
5. Any API assumption that still needs a runtime check (see the notes in `plant3d-api`).

Update `CLAUDE.md` (Known issues) and the `plant3d-api` skill whenever you learn or confirm something new.
