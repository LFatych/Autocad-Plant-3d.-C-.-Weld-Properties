#!/usr/bin/env bash
# Compile-check the real WeldPropUtils.csproj on Linux against the Plant 3D SDK reference DLLs:
# net48 (Plant 3D 2024) and net8.0-windows (Plant 3D 2025/2026).
# Usage: PLANT_REF_2024=<dir> PLANT_REF_2026=<dir> tools/compile-check/build.sh
#   each <dir> holds AcCoreMgd/AcDbMgd/AcMgd.dll and the PnP*.dll files of that SDK
#   (see .claude/skills/plant3d-build-check). Omit one variable to check only the other target.
set -euo pipefail
here="$(cd "$(dirname "$0")" && pwd)"
proj="$here/../../WeldPropUtils/WeldPropUtils/WeldPropUtils.csproj"
work="${TMPDIR:-/tmp}/weldprop-compile-check"
mkdir -p "$work"
# The csproj expects the SDK layout <root>/inc/Ac*.dll and <root>/inc-x64/PnP*.dll: fake it with symlinks.
fake_sdk() { mkdir -p "$work/sdk$1"; ln -sfn "$2" "$work/sdk$1/inc"; ln -sfn "$2" "$work/sdk$1/inc-x64"; echo "$work/sdk$1"; }
[ -n "${PLANT_REF_2024:-}" ] && export AP3D_SDK_2024="$(fake_sdk 2024 "$PLANT_REF_2024")"
[ -n "${PLANT_REF_2026:-}" ] && export AP3D_SDK_2026="$(fake_sdk 2026 "$PLANT_REF_2026")"
: "${AP3D_SDK_2024:=}${AP3D_SDK_2026:?set PLANT_REF_2024 and/or PLANT_REF_2026}"
# Linux dotnet has no WindowsDesktop SDK: skip its targets and inject the WinForms/WPF reference assemblies instead.
dotnet build "$proj" -nologo -v q --no-incremental \
  -p:ImportWindowsDesktopTargets=false \
  -p:CustomAfterMicrosoftCommonTargets="$here/LinuxWindowsDesktop.targets" \
  -p:BaseOutputPath="$work/bin/" "$@" 2>&1 \
  | grep -E "error|warning|Build succeeded|Error\(s\)|Warning\(s\)" | sed "s|$(cd "$here/../.." && pwd)/||g" | sort -u
exit "${PIPESTATUS[0]}"
