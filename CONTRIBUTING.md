# Contributing

Bug reports and pull requests are welcome. If your tracker's layout confuses
`init` or a tool, an issue with the headers and the status values (no personal
data needed) is the most useful thing you can send.

## Setup

```sh
git clone https://github.com/butyavk/tracker-mcp && cd tracker-mcp
uv sync
uv run pytest
uvx ruff check src tests evals
```

The tests build synthetic trackers (an English and a Russian layout, CSV files) and
run Google Sheets against a fake API (`tests/fake_sheets.py`), so they need no
network and no Google account. CI runs them on Linux, macOS and Windows.

To try the server by hand, the MCP Inspector:

```sh
npx @modelcontextprotocol/inspector uv --directory . run tracker-mcp --tracker /path/tracker.xlsx
```

## Changing tools

Tool descriptions are prompts: a model reads them to decide what to call. After
changing a description, an argument or the server instructions, run the evals
(`ANTHROPIC_API_KEY=… uv run python evals/run_evals.py`, a few cents with a small
model). If you add or rename a tool, update the tool list in `mcpb/manifest.json`;
a test checks it.

Every write goes through `Tracker._write` (plan, verify, backup, audit) and the
confirmation flow in `server.py`. Please keep new writes on that path.

## Releases

1. Bump the version in `pyproject.toml`, `server.json` (twice), `mcpb/manifest.json` and
   `plugin/.claude-plugin/plugin.json` (the version and the `tracker-mcp==` pin), add a
   section to `CHANGELOG.md`, run `uv lock`.
   `tests/test_packaging.py` checks that they agree.
2. Tag and push: `git tag v0.3.1 && git push origin v0.3.1`.
3. `.github/workflows/release.yml` then publishes to PyPI (trusted publishing, no
   token), to the MCP Registry (GitHub OIDC), and makes a GitHub release with the
   `.mcpb` extension attached.

`sh mcpb/build.sh` builds the extension locally (needs Node).

`plugin/` is the Claude plugin (the server plus a `weekly-review` skill), installable
from this repository's marketplace and submittable to Claude's plugin directory at
claude.ai/directory/manage. Check it with `claude plugin validate ./plugin`.
