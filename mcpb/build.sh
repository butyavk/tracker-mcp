#!/bin/sh
# Build dist/tracker-mcp-<version>.mcpb: the Claude Desktop extension (one file,
# double-click to install). Claude Desktop runs it with its own uv, which installs
# Python and the dependencies from uv.lock on first start. Needs Node for npx.
# Usage: sh mcpb/build.sh [out-dir]
set -e
cd "$(dirname "$0")/.."
OUT="${1:-dist}"
VERSION="$(sed -n 's/^version = "\(.*\)"$/\1/p' pyproject.toml | head -1)"
MCPB="npx -y @anthropic-ai/mcpb@2.1.2"
STAGE="$(mktemp -d)"
trap 'rm -rf "$STAGE"' EXIT

cp -R pyproject.toml uv.lock README.md LICENSE src examples "$STAGE"/
cp mcpb/manifest.json mcpb/icon.png "$STAGE"/
find "$STAGE" -name __pycache__ -type d -prune -exec rm -rf {} +
$MCPB validate "$STAGE/manifest.json"
mkdir -p "$OUT"
$MCPB pack "$STAGE" "$OUT/tracker-mcp-$VERSION.mcpb"
