# tracker-mcp for Claude

Connects Claude to the spreadsheet where you track your job applications (an .xlsx or
.csv file). Ask which follow-ups are due this week, which applications went silent, or
which channel brings the most replies, and tell Claude what happened ("applied to Acme
via a referral", "Globex rejected me"): it shows the exact cells it will change and
writes them only after your yes, with a backup and undo.

The plugin starts the open-source server `tracker-mcp` from PyPI with `uvx`, so you need
[uv](https://docs.astral.sh/uv/). It also adds a `weekly-review` skill. It asks for one
setting: the path to your tracker file. On first start the server writes a small config
next to the file, describing your columns and statuses; Claude tells you what to check.

What it runs: the `tracker-mcp` package, pinned to an exact version. `uv.lock` lists its
dependencies with their exact versions and hashes. What it touches: your tracker file,
its config, and a `backups-tracker-mcp` folder next to it for backups and an audit log.
It sends nothing anywhere and has no telemetry.

Install in Claude Code:

    /plugin install tracker-mcp --marketplace butyavk/tracker-mcp

Full documentation, Google Sheets setup and the security model:
https://github.com/butyavk/tracker-mcp
