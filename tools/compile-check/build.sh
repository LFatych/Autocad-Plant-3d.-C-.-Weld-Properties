#!/usr/bin/env bash
# Compile-check the real WeldPropUtils.csproj (Plant 3D 2026, net8.0-windows) on Linux against the SDK reference DLLs.
# Usage: PLANT_REF_2026=<dir> tools/compile-check/build.sh
#   <dir> holds AcCoreMgd/AcDbMgd/AcMgd.dll and the PnP*.dll files of the 2026 SDK (see .claude/skills/plant3d-build-check).
set -euo pipefail
here="$(cd "$(dirname "$0")" && pwd)"
proj="$here/../../WeldPropUtils/WeldPropUtils/WeldPropUtils.csproj"
: "${PLANT_REF_2026:?set PLANT_REF_2026 to the folder with the 2026 SDK reference DLLs}"
work="${TMPDIR:-/tmp}/weldprop-compile-check"
# The csproj expects the SDK layout <root>/inc/Ac*.dll and <root>/inc-x64/PnP*.dll: fake it with symlinks.
mkdir -p "$work/sdk2026"
ln -sfn "$PLANT_REF_2026" "$work/sdk2026/inc"
ln -sfn "$PLANT_REF_2026" "$work/sdk2026/inc-x64"
export AP3D_SDK_2026="$work/sdk2026"
# Linux dotnet has no WindowsDesktop SDK: skip its targets and inject the WinForms/WPF reference assemblies instead.
dotnet build "$proj" -nologo -v q --no-incremental \
  -p:ImportWindowsDesktopTargets=false \
  -p:CustomAfterMicrosoftCommonTargets="$here/LinuxWindowsDesktop.targets" \
  -p:BaseOutputPath="$work/bin/" "$@" 2>&1 \
  | grep -E "error|warning|Build succeeded|Error\(s\)|Warning\(s\)" | sed "s|$(cd "$here/../.." && pwd)/||g" | sort -u
exit "${PIPESTATUS[0]}"
