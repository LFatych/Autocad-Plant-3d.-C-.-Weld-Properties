#!/usr/bin/env bash
# Compile-check WeldPropUtils (C# and XAML) on Linux against the Plant 3D SDK reference DLLs.
# Usage: PLANT_REF_2027=<dir> [PLANT_REF_2026=<dir>] tools/compile-check/build.sh
#   Each <dir> holds AcCoreMgd/AcDbMgd/AcMgd.dll and the PnP*.dll files of that SDK
#   (see .claude/skills/plant3d-build-check). Every version whose variable is set is checked.
#   2027 needs the .NET 10 SDK, 2026 the .NET 8 SDK.
set -uo pipefail
here="$(cd "$(dirname "$0")" && pwd)"
repo="$(cd "$here/../.." && pwd)"
work="${TMPDIR:-/tmp}/weldprop-compile-check"
mkdir -p "$work"

# XAML compiler (PresentationBuildTasks) from NuGet, once.
wpf="$work/wpfsdk"
if [ ! -f "$wpf/targets/Microsoft.WinFx.targets" ]; then
  curl -sSL -o "$work/wpfsdk.nupkg" https://api.nuget.org/v3-flatcontainer/microsoft.net.sdk.windowsdesktop/3.0.0/microsoft.net.sdk.windowsdesktop.3.0.0.nupkg
  rm -rf "$wpf"; mkdir -p "$wpf"; (cd "$wpf" && unzip -q "$work/wpfsdk.nupkg")
fi

status=0
checked=0
for version in 2027 2026; do
  refvar="PLANT_REF_$version"; ref="${!refvar:-}"
  [ -z "$ref" ] && continue
  checked=1
  case $version in 2027) tfm=net10.0-windows;; 2026) tfm=net8.0-windows;; esac
  # Build a copy: the markup compiler writes generated files next to the XAML sources.
  src="$work/src$version/WeldPropUtils"
  rm -rf "$work/src$version"; mkdir -p "$src"
  (cd "$repo/WeldPropUtils/WeldPropUtils" && tar --exclude=./bin --exclude=./obj -cf - .) | (cd "$src" && tar -xf -)
  # The markup compiler joins paths with '\': make "<dir>\<entry>" and "<obj>/<tfm>/\" resolve.
  mkdir -p "$src/obj/Debug/$tfm"
  for entry in "$src"/*; do ln -sfn "$entry" "$src\\$(basename "$entry")"; done
  ln -sfn / "$src/obj/Debug/$tfm/\\"
  # Pass 1 writes "<obj>\\WeldPropUtils_MarkupCompile.lref"; pass 2 (XAML using our own controls) only runs
  # when "<obj>/WeldPropUtils_MarkupCompile.lref" exists.
  ln -sfn "\\WeldPropUtils_MarkupCompile.lref" "$src/obj/Debug/$tfm/WeldPropUtils_MarkupCompile.lref"
  echo "== Plant 3D $version ($tfm)"
  dotnet build "$src/WeldPropUtils.csproj" -nologo -v q -c Debug \
    -p:PlantVersion=$version -p:AcadDir="$ref" -p:WpfSdkDir="$wpf" \
    -p:ImportWindowsDesktopTargets=false \
    -p:CustomAfterMicrosoftCommonTargets="$here/LinuxWindowsDesktop.targets" "$@" 2>&1 \
    | grep -E "error|warning|Build succeeded|Error\(s\)|Warning\(s\)" | sed "s|$work/src$version/|WeldPropUtils/|g" | sort -u
  [ "${PIPESTATUS[0]}" -eq 0 ] || status=1
done
[ $checked -eq 1 ] || { echo "set PLANT_REF_2027 and/or PLANT_REF_2026"; exit 2; }
exit $status
