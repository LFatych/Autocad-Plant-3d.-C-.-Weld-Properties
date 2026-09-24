#!/usr/bin/env bash
# Installs the tools Claude needs for the compile check (see .claude/skills/plant3d-build-check).
# Only runs in Claude Code cloud sessions; does nothing on a developer machine.
[ "${CLAUDE_CODE_REMOTE:-}" = "true" ] || exit 0
if ! command -v dotnet >/dev/null 2>&1; then
  (apt-get update -q && apt-get install -y -q dotnet-sdk-8.0) >/tmp/dotnet-install.log 2>&1 \
    || echo "dotnet-sdk-8.0 install failed, see /tmp/dotnet-install.log" >&2
fi
exit 0
