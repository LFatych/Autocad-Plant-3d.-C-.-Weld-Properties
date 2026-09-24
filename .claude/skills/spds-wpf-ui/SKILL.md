---
name: spds-wpf-ui
description: SPDS company look for the plugin's WPF windows — the approved dark and light themes (exact brand colours from the SPDS logo), fonts, logo, theme switching, how to host a WPF window inside AutoCAD Plant 3D, and the approved Weld Property Mapping window layout. Use whenever creating or changing any WPF/XAML UI, dialog, colour or icon in this repo.
---

# SPDS WPF UI

Approved concept (clickable mockup): https://claude.ai/artifact/N2pdBr276TeyNWUGSzb119 (private to the user).
Decision (user, 2026-09-24): **two themes, dark (default) and light, switchable in the window**. The earlier
"graphite frame + light workspace" option was dropped.

## Brand
| Token | Hex | Source |
|---|---|---|
| SPDS yellow | `#FDD541` | sampled from the SPDS logo |
| SPDS graphite (blue-grey slate) | `#44525C` | sampled from the SPDS logo |
| Pastel yellow (tint) | `#FFF1BF` | derived; light-theme highlights |

- Logo: `ppt/media/image3.png` inside `SPDS_Plant3D_Python_Course_Agenda.pptx` on the user's Drive (file ID
  `1ABgBrRRmJtBm9Y7wxj3TxO2R9-QPMmuI`). Crop it to its alpha bounds, make it square, and ship 32/48/64 px PNGs as WPF `Resource`s. Ask the
  user before committing the logo to a public repo.
- Fonts: headings **Century Gothic** (the SPDS slides use it; it ships with Windows/Office), falling back to `Segoe UI`.
  Body text **Segoe UI** 12–13 px. Property/field names **Consolas**.

## Theme resource keys (both dictionaries define every key)
`Themes/SpdsDark.xaml` and `Themes/SpdsLight.xaml`, as `SolidColorBrush` resources named `Spds.<Key>`:

| Key | Dark | Light | Used for |
|---|---|---|---|
| WindowBg | `#1F262B` | `#F5F4EF` | window background |
| ChromeBg | `#29323A` | `#FFFFFF` | header + footer bars |
| ChromeText / ChromeMuted | `#EEF0EE` / `#A9B3BA` | `#2E373E` / `#5A6770` | text in bars |
| HeaderRule | `#FDD541` (2 px) | `#FDD541` (3 px) | line under header |
| PanelBg | `#232B31` | `#EFEEE8` | left/right side panels |
| CardBg | `#29323A` | `#FFFFFF` | grouped sections |
| FieldBg | `#313C45` | `#FFFFFF` | inputs, combos |
| Border / BorderStrong | `#3E4A54` / `#52606B` | `#D9DCDD` / `#B3BABF` | outlines |
| Text / Muted / Faint | `#EEF0EE` / `#A9B3BA` / `#8F9AA2` | `#2E373E` / `#5A6770` / `#5E6B74` | text levels |
| Mono | `#D3D9DD` | `#384650` | Consolas field names |
| ChipBg / ChipBorder | `#2B343B` / `#3E4A54` | `#FFFFFF` / `#D9DCDD` | draggable property items |
| SelectedBg / SelectedBorder | `#3A3A2C` / `#FDD541` | `#FFF1BF` / `#C99A12` | selected property |
| DropMappedBg / DropEmptyBg | `#313C45` / `#20282D` | `#F7F6F1` / `#FFFFFF` | drop targets |
| DropEmptyBorder (dashed) | `#5A6772` | `#AEB6BB` | empty drop target |
| DropOverBg / DropOverBorder | `#3A3A2C` / `#FDD541` | `#FFF1BF` / `#44525C` | drag hover |
| TabOnBg / TabOnText | `#FDD541` / `#1F262B` | `#44525C` / `#FFFFFF` | active filter tab |
| Side1 (+ ring) | `#FDD541` | `#FDD541` (ring `#B8900F`) | port 1 / larger side |
| Side2 | `#9FB3C2` | `#44525C` | port 2 / smaller side |
| PrimaryBg / PrimaryText | `#FDD541` / `#1F262B` | `#FDD541` / `#2E373E` (border `#E0B21F`) | main action button |
| NoteBg / NoteBorder / NoteText / NoteIcon | `#2C3740` / `#465460` / `#DCE2E5` / `#9FB3C2` | `#FFF6D6` / `#EBCF6A` / `#4A3E12` / `#9C7A0C` | info box |
| Unmapped | `#F2A07B` | `#B4532A` | "not mapped" values |
| CheckAccent | `#FDD541` | `#44525C` | checkbox/radio accent |

Text colours meet 4.5:1 contrast on their backgrounds. Keep that true when you add colours. Side 1 and side 2 must
differ in lightness, not only in hue.

## Theme mechanics in WPF (inside AutoCAD)
- **Never touch `System.Windows.Application.Current.Resources`.** AutoCAD owns that application. Put the theme
  dictionary in the window's own `Resources.MergedDictionaries`, and swap that entry to switch themes.
- Controls use `{DynamicResource Spds.X}` so a theme switch applies at once.
- Default theme: follow AutoCAD's `COLORTHEME` system variable (0 = dark, 1 = light), read with
  `Application.GetSystemVariable("COLORTHEME")` (verified in `AcCoreMgd`). The user's own choice overrides it and is stored in the
  per-user settings.
- Dark title bar on Windows 10/11: optional, via `DwmSetWindowAttribute(DWMWA_USE_IMMERSIVE_DARK_MODE=20)`.

## Hosting the window (verified in AcCoreMgd 2024 + 2026)
`Autodesk.AutoCAD.ApplicationServices.Core.Application` (the AcMgd `Application` inherits it):
```
bool? ShowModalWindow(Window formToShow)
bool? ShowModalWindow(Window owner, Window formToShow, bool persistSizeAndPosition)
bool? ShowModalWindow(IntPtr owner, Window formToShow, bool persistSizeAndPosition)
void  ShowModelessWindow(Window formToShow)            // + the same owner/persist overloads
```
Use `Application.ShowModalWindow(window)`, not `window.ShowDialog()`. It parents the window to AutoCAD, handles focus,
and can remember its size and position. A modal window runs inside the command, so Plant/DWG access afterwards stays in
document context.

## Project/build notes
- Add `<UseWPF>true</UseWPF>` to the SDK-style csproj. It covers both targets, and XAML files are picked up as `Page` items automatically.
- **XAML cannot be compile-checked on Linux.** The WPF markup compiler only ships with the Windows SDK. So:
  - keep the logic (the mapping model, profile load/save, reading the schema) in plain C# classes, which `tools/compile-check` does verify
  - keep XAML declarative (bindings plus a thin code-behind)
  - tell the user that the XAML is only verified by their Visual Studio build.
- Drag and drop: WPF `DragDrop.DoDragDrop` with the property id as data. Also support click-to-assign (select, then click a
  target) and keyboard use (Enter on a focused target), as in the mockup.

## Mapping window layout (approved)
- **Header:** logo, title "Weld Property Mapping", subtitle with the project name, theme toggle, profile combo, Import/Export.
- **Left panel:** connected-part properties read from Project Setup, with search and class tabs (All/Pipe/Fittings/Flanges/Valves/Nozzle).
- **Centre:** weld class combo and a "Same mapping for both sides" check. Side 1 and side 2 groups, each with 5 drop targets
  (Material, OD, WallThickness, LDS, SPEC + side number). "Add a weld property from Project Setup". WeldNumber is not mapped.
- **Right panel:** preview for a picked weld (both part cards and the values to write) plus an info note on unmapped fields.
- **Footer:** hint text, "Apply to: All welds in drawing / Selected welds", Cancel, and the primary "Apply mapping" button.
