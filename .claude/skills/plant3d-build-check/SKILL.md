---
name: plant3d-build-check
description: Compile WeldPropUtils on Linux against the real Plant 3D 2024 (net48) and 2026 (net8.0-windows) SDK reference DLLs fetched from the user's Google Drive, to catch compile errors and API mistakes before pushing. Use before every commit that touches .cs or project files, and whenever you need to prove an API member exists in a given Plant 3D version.
---

# Compile check against the real SDKs

This only proves the code compiles. It does not run it, because running needs Windows + Plant 3D. The user still tests in Plant 3D.

## 1. Tools (the SessionStart hook normally does this)
```bash
command -v dotnet || (apt-get update -q && apt-get install -y -q dotnet-sdk-8.0)
```
NuGet (api.nuget.org) is reachable. The check project downloads `Microsoft.NETFramework.ReferenceAssemblies`
(for net48) and `Microsoft.WindowsDesktop.App.Ref` (WinForms/WPF refs for net8 on Linux).

## 2. Reference DLLs (Google Drive MCP, once per session)
Download each file with `mcp__Google_Drive__download_file_content`. Large results are saved to a JSON file
(`{content: base64, title, id}`). Decode it into a folder:
```bash
jq -r .content "$saved_json" | base64 -d > "$REF/$(jq -r .title "$saved_json")"
```
Use `$SCRATCH/sdk2024/ref` and `$SCRATCH/sdk2026/ref` (scratchpad). **Never commit these DLLs** (Autodesk license).

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
- `plantsdk_dev.chm`: 2024 `1UeabGpxKoq6yB4vEHvug9VeK1-2xw7qy`

Extract them with `7z x` (install with `apt-get install -y p7zip-full`).
If a file ID stops working, search the Drive with `title = 'PnPDataLinks.dll'` and check the parent folder.
If the Drive tools return "Insufficient scope", ask the user to reconnect Google Drive with read access.

## 3. Run
```bash
PLANT_REF_2024=$SCRATCH/sdk2024/ref PLANT_REF_2026=$SCRATCH/sdk2026/ref TMPDIR=$SCRATCH \
  tools/compile-check/build.sh
```
- `build.sh` builds the **real** `WeldPropUtils.csproj`. It points `AP3D_SDK_2024/2026` at symlinked fake SDK folders,
  skips the WindowsDesktop targets (they don't exist on Linux), and injects the WinForms/WPF reference assemblies through
  `tools/compile-check/LinuxWindowsDesktop.targets`. Set only one `PLANT_REF_*` variable to check a single target.
- Any `error` fails the check. Fix it before committing.
- A new `warning CS…` in code you touched counts as a finding: fix it or explain it.
- The known baseline warning is `CS0642` in `WeldPropertiesHandler.cs` (the `using (DocumentLock …) ;` bug), until that bug is fixed.
- New `.cs` files are picked up automatically (SDK-style project). XAML is not compiled on Linux (see `spds-wpf-ui`).
- Plant 3D 2027 / .NET 10: once an SDK exists, add a `net10.0-windows` target and a `PLANT_REF_2027` variable, the same way.

## 4. Report
Tell the user which targets compiled and quote any warnings. Say plainly that runtime behaviour is not tested yet.
