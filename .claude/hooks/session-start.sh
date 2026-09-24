#!/usr/bin/env bash
# Installs the tools Claude needs for the compile check (see .claude/skills/plant3d-build-check).
# Only runs in Claude Code cloud sessions; does nothing on a developer machine.
# .NET 10 SDK builds the Plant 3D 2027 target (default); .NET 8 SDK the optional 2026 target.
[ "${CLAUDE_CODE_REMOTE:-}" = "true" ] || exit 0
missing=""
dotnet --list-sdks 2>/dev/null | grep -q '^10\.' || missing="$missing dotnet-sdk-10.0"
dotnet --list-sdks 2>/dev/null | grep -q '^8\.' || missing="$missing dotnet-sdk-8.0"
if [ -n "$missing" ]; then
  (apt-get update -q && apt-get install -y -q $missing) >/tmp/dotnet-install.log 2>&1 \
    || echo "$missing install failed, see /tmp/dotnet-install.log" >&2
fi
exit 0
