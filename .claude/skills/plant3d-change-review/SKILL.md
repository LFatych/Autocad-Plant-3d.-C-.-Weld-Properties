---
name: plant3d-change-review
description: Checklist to run before committing any change to this Plant 3D plugin, since the code cannot be compiled or run in Claude's Linux environment. Use before every commit, when reviewing a diff, or when the user asks "is this ready?".
---

# Pre-commit review for WeldPropUtils

This project cannot be built here, because there are no Autodesk DLLs and no .NET Framework. Review by hand
instead. Go through every point below, and fix problems before you commit.

## 1. Compiles in principle
- [ ] Every type or member you used already appears in the repo, in the `plant3d-api` skill, or in the SDK.
      If you are unsure, say so in the summary. Do not guess.
- [ ] Only C# 7.3 syntax: no `using var`, no switch expressions, no `??=`, no `^1` or `..`, no records.
- [ ] Each new `.cs` file has a `<Compile Include=...>` line in `WeldPropUtils.csproj` (old-style csproj!).
- [ ] Each new Autodesk reference uses an `$(AP3D_SDK_2024)` HintPath and `<Private>False</Private>`.
- [ ] `using` directives and aliases (`pPart`, `pPort`, `portCol`) resolve without ambiguity.

## 2. Runtime correctness
- [ ] Document, Database, Editor and DataLinksManager are fetched when the command runs.
- [ ] Transactions are committed. No DBObject is used after its transaction ends.
- [ ] Null checks exist for `as` casts, `FindAcPpRowId` results (≤ 0) and missing dictionary keys.
- [ ] Numbers are parsed with `CultureInfo.InvariantCulture`.
- [ ] No `MessageBox` inside loops. Errors go to `ed.WriteMessage`.
- [ ] Every write goes through `dlm.SetProperties`, unless there is a reason to write rows directly.

## 3. Behaviour
- [ ] Existing commands keep their names and outputs unless the user asked to change them.
- [ ] Any new custom property that must exist in Project Setup is listed in the summary and in `CLAUDE.md`.

## 4. Hand-off to the user
In the final message, include:
1. What changed and why (short).
2. **Build**: "Build the Release configuration in Visual Studio, then NETLOAD `bin\Release\WeldPropUtils.dll`."
3. **Manual test plan**: which kind of drawing to open (butt weld, tap, socket weld, weld to an equipment nozzle,
   weld with only one side connected), which command to run, and which property values to check in the Properties
   palette or Data Manager.
4. Any API assumption that still needs checking against the SDK.

Update `CLAUDE.md` (Known issues) and the `plant3d-api` skill when you learn or confirm something new.
