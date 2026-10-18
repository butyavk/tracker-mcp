#!/bin/sh
# Launcher for Claude Desktop: GUI apps don't get the shell PATH, so look for uv
# in the usual install locations, then start the server over stdio.
export PATH="$HOME/.local/bin:$HOME/.cargo/bin:$HOME/.homebrew/bin:/opt/homebrew/bin:/usr/local/bin:$PATH"
cd "$(dirname "$0")" || exit 1
exec uv run --quiet --frozen tracker-mcp
