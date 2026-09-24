---
name: plant3d-build-check
description: Compile WeldPropUtils on Linux (C# and XAML) against the real Plant 3D 2027 (net10.0-windows) SDK reference DLLs, and optionally 2026 (net8.0-windows), fetched from the user's Google Drive, to catch compile errors and API mistakes before pushing. Use before every commit that touches .cs or project files, and whenever you need to prove an API member exists in a given Plant 3D version.
---

# Compile check against the real SDKs

This only proves the code compiles. It does not run it, because running needs Windows + Plant 3D. The user still tests in Plant 3D.

## 1. Tools (the SessionStart hook normally does this)
```bash
dotnet --list-sdks | grep -q '^10\.' || (apt-get update -q && apt-get install -y -q dotnet-sdk-10.0)
dotnet --list-sdks | grep -q '^8\.'  || apt-get install -y -q dotnet-sdk-8.0   # only for the 2026 check
```
NuGet (api.nuget.org) is reachable. The build downloads `Microsoft.WindowsDesktop.App.Ref` (WinForms/WPF reference
assemblies: 10.0.12 for net10, 8.0.0 for net8) and, once, `Microsoft.NET.Sdk.WindowsDesktop` 3.0.0 for the XAML compiler.

## 2. Reference DLLs (Google Drive MCP, once per session)
Download each file with `mcp__Google_Drive__download_file_content`. Large results are saved to a JSON file
(`{content: base64, title, id}`). Decode it into a folder:
```bash
jq -r .content "$saved_json" | base64 -d > "$REF/$(jq -r .title "$saved_json")"
```
Use `$SCRATCH/sdk2027/ref` (and `$SCRATCH/sdk2026/ref`) in the scratchpad. 2027 is the default target; 2026 is still
buildable with `PlantVersion=2026`; 2024 IDs are only for API research. **Never commit these DLLs** (Autodesk license).

2027 (`AP3D_2027_SDK`, folder `1bHCL2xQfAR9JbWepFa0TAloD7PHH1boo`):

| File | 2027 file ID |
|---|---|
| AcCoreMgd.dll | `1n9qzKjAIkRvyHiiBdXX9ng42POLU_1B3` |
| AcDbMgd.dll | `14g9kQG0_WwOezxCLBgzkyMWgmWLm04iP` |
| AcMgd.dll | `1xLbqJRQGquzjG8LfqXBzxEIwF1gK0l4p` |
| PnP3dObjectsMgd.dll | `1VH4ARlWxWwvGCwWAAC-ZWFvU-2DfQGbE` |
| PnP3dProjectPartsMgd.dll | `15N9agU_enmzh_AVqy1ktxXYwAkMlB7tU` |
| PnPDataLinks.dll | `1UOlno5OoZ6GhUuCPvj8gUhDaBLJF1KHR` |
| PnPDataObjects.dll | `1TKoozzDvf1C7siWSSPcbmXYH4wOK5rM3` |
| PnPProjectManagerMgd.dll | `1v82g9O6doJyncsE_eeyRbOMwms7Hi1lT` |
| PnPCommonMgd.dll | `1u_8IClqsoOisv4j9RrlwW6V-7j1T_0Kr` |
| PnP3dStructureObjectsMgd.dll (new in 2027, not referenced yet) | `189F0NoNeXneutXf-WEq2p_OM9yefGmBb` |

Older SDKs:

| File | 2024 file ID (`AP3D_2024_SDK`) | 2026 file ID (`AP3D_2026_SDK`) |
|---|---|---|
| AcCoreMgd.dll (`inc`) | `1Ix3SMf2eKh3nh5ma5mWb1Bw3ogsXy-3P` | `1AsakjMsxrQyzXtEQnvWpkUU6lbWWgBFi` |
| AcDbMgd.dll (`inc`) | `1pTwXBx1xhzeKe_JRQMT5yNW8n8sb-S5C` | `1GFvcfzebxqYXDa2RiWxdDFT63EZCif5c` |
| AcMgd.dll (`inc`) | `1gKGSwNRtDqGBUFTTu8NENrGy1kx4ksdu` | `1Z1y_1yJfTG7bgHYTspFfBHjd8dfFNh0x` |
| PnP3dObjectsMgd.dll | `1ExdFQZ7hhx5jJkdqneWmQhqLfZFK3Kf7` | `1hdnjtJShl_Czfh8x0mGQh2p4psjdjN26` |
| PnP3dProjectPartsMgd.dll | `1OLf9-5TKvjv5woyDliPn34KP14kefNpd` | `1RvjsN19ACvwIURFrjOMFGQl6B5disWr4` |
| PnPDataLinks.dll | `1aUfl6sv15PVqVRlkYbLHl3mDc3-LJpHa` | `1Zaf3IFfxAwaUqUqQy-ebE-vQdFWhoXxU` |
| PnPDataObjects.dll | `1Nvmmf5nYI9drGGmBCvQeUVD-3BUiSssk` | `1Spgkz99yZs06esu-77Y4Ov3ZbaHaSSHm` |
| PnPProjectManagerMgd.dll | `1t3z2omdVK944RWaAkhYITWUad-shTbIR` | `1ilme4UtZNE12fYM0vEF4aFPjAmL1WtWF` |
| PnPCommonMgd.dll | `1VnapVN2f4RfLKLClyOruPVUTte90KdN3` | `1JodI_GxEzweZj7SYj4in5BMvqiY97wNO` |

Help files, which are optional and useful for API research:
- `plantsdk_ref.chm`: 2024 `15VDvnLXmR4NH6wO7hu5Pqu9zaRuhOE3t`, 2026 `1_3yDFSuGNaSHk2a0t3BlT00_vMzJVwNd`
- `plantsdk_dev.chm`: 2024 `1UeabGpxKoq6yB4vEHvug9VeK1-2xw7qy`, 2027 `1OzQDGBN6Ua709d6rCz4GpvDDUycftkVD`
- `plantsdk_ref.chm` 2027: `1VJODY1mvBkgoiLBtTYdhQu07qs4AQgiC`

Extract them with `7z x` (install with `apt-get install -y p7zip-full`). To list the API of a DLL, decompile it with
`ilspycmd` **11.x** (`dotnet tool install -g ilspycmd`; run with `DOTNET_ROLL_FORWARD=Major`); 8.x fails on the net10 DLLs.
Never commit decompiled output.
If a file ID stops working, search the Drive with `title = 'PnPDataLinks.dll'` and check the parent folder.
If the Drive tools return "Insufficient scope", ask the user to reconnect Google Drive with read access.

## 3. Run
```bash
PLANT_REF_2027=$SCRATCH/sdk2027/ref PLANT_REF_2026=$SCRATCH/sdk2026/ref TMPDIR=$SCRATCH \
  tools/compile-check/build.sh
```
- Every version whose `PLANT_REF_<ver>` is set is checked. Always check 2027; check 2026 too while it is kept buildable.
- `build.sh` copies the project to `$TMPDIR/weldprop-compile-check/src<ver>` and builds the **real** `WeldPropUtils.csproj`
  with `-p:PlantVersion=<ver> -p:AcadDir=<ref dir>` (the csproj resolves references from `AcadDir` and `AcadDir/PLNT3D`).
  It skips the WindowsDesktop targets (they don't exist on Linux) and, through `tools/compile-check/LinuxWindowsDesktop.targets`,
  injects the WinForms/WPF reference assemblies and the WPF markup compiler (`Microsoft.WinFx.targets` from NuGet).
- **XAML is compiled** (both passes). Errors show as `MCxxxx` with file/line, e.g. MC3074 unknown tag, MC3072 unknown
  property (also on our own controls, pass 2), CS1061 for a missing code-behind handler, CS0103 for a wrong `x:Name`.
  The markup compiler joins paths with `\`; `build.sh` makes those resolve with symlinks (`<src>\<entry>`, `obj/.../\`,
  and `WeldPropUtils_MarkupCompile.lref`). BAML is not embedded in the Linux DLL, so that DLL is never shipped.
  Not covered: runtime-only problems (a `StaticResource` key that doesn't exist, a wrong pack URI, binding paths).
- Any `error` fails the check. Fix it before committing.
- A new `warning CS…` in code you touched counts as a finding: fix it or explain it.
- There are no baseline warnings (the old `CS0642` was fixed). Any warning is new.
- New `.cs` and `.xaml` files are picked up automatically (SDK-style project; `*.xaml` → `Page`).

## 4. Report
Tell the user whether it compiled and quote any warnings. Say plainly that runtime behaviour is not tested yet.
