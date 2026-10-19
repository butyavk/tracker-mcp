#!/bin/sh
# One-time setup on macOS: install dependencies, run the tests, build the
# launcher app for Claude Desktop, and register the server in Claude Desktop's
# config (the old config is backed up first).
# Usage: sh install.sh /absolute/path/to/tracker.xlsx
set -e
if [ "$(uname -s)" != "Darwin" ]; then
  echo "install.sh is for Claude Desktop on macOS. On Linux and Windows, add the server to your"
  echo "client's config as the README shows (Quick start)."
  exit 1
fi
cd "$(dirname "$0")"
DIR="$(pwd)"
export PATH="$HOME/.local/bin:$HOME/.cargo/bin:$HOME/.homebrew/bin:/opt/homebrew/bin:/usr/local/bin:$PATH"

if ! command -v uv >/dev/null 2>&1; then
  echo "uv is not installed. Install it, then run this script again:"
  echo "  curl -LsSf https://astral.sh/uv/install.sh | sh     (or: brew install uv)"
  exit 1
fi
if [ ! -f "$1" ]; then
  echo "Usage: sh install.sh /absolute/path/to/tracker.xlsx"
  exit 1
fi
TRACKER="$(cd "$(dirname "$1")" && pwd)/$(basename "$1")"

echo "== Dependencies"
uv sync --frozen --quiet
echo "== Tests"
uv run --frozen pytest -q -p no:cacheprovider

echo "== Config"
if ! TRACKER_PATH="$TRACKER" uv run --frozen tracker-mcp --check >/dev/null; then
  echo "Fix the config (every key is described in examples/tracker-mcp.example.toml), then run again."
  exit 1
fi
echo "OK: ${TRACKER_CONFIG:-$(dirname "$TRACKER")/$(basename "$TRACKER" .xlsx).tracker-mcp.toml}"

# Claude Desktop starts local servers without its own privacy permissions, so a
# bare /bin/sh cannot read ~/Documents. The launcher app owns that permission
# instead (see macos/launcher.c). It is rebuilt only when its sources change,
# because macOS ties the permission to this exact build.
echo "== Launcher app"
APP="$HOME/Applications/Tracker MCP.app"
BIN="$APP/Contents/MacOS/tracker-mcp"
if [ -x "$BIN" ] && [ "$BIN" -nt macos/launcher.c ] && [ "$BIN" -nt macos/Info.plist ]; then
  echo "Up to date: $APP"
else
  if ! command -v cc >/dev/null 2>&1; then
    echo "No C compiler. Install the Command Line Tools (xcode-select --install), then run this script again."
    exit 1
  fi
  mkdir -p "$APP/Contents/MacOS"
  cp macos/Info.plist "$APP/Contents/Info.plist"
  cc -O2 -Wall -o "$BIN" macos/launcher.c
  codesign --force --sign - "$APP"
  echo "Built: $APP"
fi

echo "== Claude Desktop config"
if pgrep -x Claude >/dev/null 2>&1; then
  echo "Claude is running, and it rewrites its config when it quits."
  echo "Quit Claude (Cmd+Q), then run this script again."
  exit 1
fi
CONFIG="$HOME/Library/Application Support/Claude/claude_desktop_config.json"
uv run --frozen python - "$CONFIG" "$BIN" "$DIR/run.sh" "$TRACKER" <<'EOF'
import json, shutil, sys, time
from pathlib import Path

config, launcher, run_sh, tracker = Path(sys.argv[1]), sys.argv[2], sys.argv[3], sys.argv[4]
data = {}
if config.exists():
    backup = config.with_name(f"{config.name}.backup-{time.strftime('%Y%m%d-%H%M%S')}")
    shutil.copy2(config, backup)
    print(f"Backup: {backup}")
    data = json.loads(config.read_text(encoding="utf-8") or "{}")
data.setdefault("mcpServers", {})["tracker"] = {
    "command": launcher,
    "args": ["/bin/sh", run_sh],
    "env": {"TRACKER_PATH": tracker},
}
config.parent.mkdir(parents=True, exist_ok=True)
config.write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
print(f"Added server 'tracker' to {config}")
EOF

echo
echo "Done. Open Claude. On first start macOS asks whether \"Tracker MCP\" may access"
echo "your Documents folder: allow it. Then check Settings -> Developer -> tracker."
echo "For Claude Code (runs with Terminal's permissions, no launcher needed):"
echo "  claude mcp add tracker -e TRACKER_PATH=\"$TRACKER\" -- /bin/sh \"$DIR/run.sh\""
