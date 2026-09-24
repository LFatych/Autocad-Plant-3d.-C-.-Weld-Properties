#!/usr/bin/env bash
# Compile-check WeldPropUtils against Plant 3D 2024 (net48) and 2026 (net8.0-windows) SDK DLLs.
# Usage: PLANT_REF_2024=<dir> PLANT_REF_2026=<dir> tools/compile-check/build.sh
# See .claude/skills/plant3d-build-check/SKILL.md for how to obtain the reference DLLs.
set -euo pipefail
here="$(cd "$(dirname "$0")" && pwd)"
: "${PLANT_REF_2024:?set PLANT_REF_2024 to the folder with the 2024 SDK reference DLLs}"
: "${PLANT_REF_2026:?set PLANT_REF_2026 to the folder with the 2026 SDK reference DLLs}"
out="${TMPDIR:-/tmp}/weldprop-compile-check"
dotnet build "$here/CompileCheck.csproj" -nologo -v q --no-incremental -o "$out" "$@" 2>&1 \
  | grep -E "error|warning|Build succeeded|Error\(s\)|Warning\(s\)" | sort -u
exit "${PIPESTATUS[0]}"
