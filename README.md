# tracker-mcp

A small MCP server over my job-search tracker. The tracker is an Excel file with
about a hundred applications in it. Instead of opening the file, I ask Claude
"which follow-ups are due this week?", "what's the status with Acme?" or
"how many did I submit in October?". Claude can also record what happened:
"mark row 12 as applied today via Ashby", "Globex rejected me".

The file stays the source of truth: I still open it in Excel, and other scripts
edit it too. Most of the code is there to make writes safe in that setting.

Nothing about my particular tracker is in the code. The sheet, the column
headers, the status words, the follow-up rule and the texts the server appends
all come from a small TOML file next to the workbook (see [Config](#config)).

## Tools

| Tool | Kind | What it does |
|---|---|---|
| `describe_tracker` | read | How this workbook is laid out: fields → columns, allowed and closed statuses, follow-up rule |
| `list_applications` | read | Filter by status, company/role text, submit dates |
| `get_application` | read | Every column of one application (by text or row) |
| `followups_due` | read | Follow-ups due within N days, overdue first; submitted rows with no follow-up date |
| `stale_applications` | read | Still waiting, submitted N+ days ago, no follow-up planned ahead |
| `summary` | read | Counts by status and month; a cohort for a date range |
| `recent_writes` | read | Audit log of this server's writes |
| `add_application` | write | New row with the not-applied status; other columns via `extra`; refuses duplicates |
| `mark_applied` | write | Submitted: date, channel, status, follow-up date by the configured rule |
| `set_status` | write | New status from the file's dropdown; a closed status cancels the pending follow-up |
| `log_followup` | write | Log a sent follow-up, schedule the next one or close |
| `close_stale` | write | Close all stale applications (or chosen rows) in one write, one backup, one undo step |
| `undo_last_write` | write | Restore the backup of the last write if nothing changed since |

Read tools are annotated `readOnlyHint`; writes carry `destructiveHint` where they
overwrite cells. All write tools take `dry_run`, `expected_version` and `confirm`.

## How a write stays safe

1. **Lock file check.** If Excel or LibreOffice has the file open (`~$…`,
   `.~lock.…#`), the write is refused. Otherwise Excel's next save would overwrite it.
2. **Optimistic concurrency.** Each read returns `version`: a hash of the file's bytes.
   A write with `expected_version` is refused if the file changed since. It works
   like HTTP `If-Match`.
3. **The user confirms, and the server enforces it.** A "please ask first" line in
   the instructions only works if the model follows it, so the yes is checked on the
   server. Before the tool body runs, a resolver computes the plan, the exact cell
   changes (field, old → new) at the current file version:
   - If the client can show a dialog (MCP elicitation), the server asks the user
     there, with the plan as the message. Older protocol versions get a request in
     the middle of the call; 2026-07-28+ gets an input-required round. A decline
     writes nothing.
   - If the client has no dialog, the write needs `confirm`: the token from a dry
     run of the same call. The token fingerprints the changes and the version, so a
     token from another preview, or from before the file changed, does not match.

   Either way the body writes with the plan's version as If-Match, so what lands in
   the file is exactly what was approved, or nothing. `[safety] confirm_writes =
   false` turns this off.
4. **Write, verify, swap.** Save to a temp file, re-open it and check every intended
   cell. Check the version once more, back up the current file, then `os.replace`,
   which is atomic.
5. **Audit and undo.** One JSON line per write: tool, row, changes, backup file.
   Undo works only if the file still matches the state right after that write, so
   it never wipes someone else's edit.

Validation comes from the workbook itself. Allowed statuses come from the column's
dropdown. Columns are found by header text, not letter. When a row is added, the
server extends the dropdowns, the autofilter and the summary sheet's formulas to cover it.

Expected refusals (a conflict, an unknown status, a duplicate) reach the model as a
readable tool error (`ToolError`). Any other exception counts as a crash, and the
SDK keeps its text out of the response.

## Config

`my-tracker.xlsx` is described by `my-tracker.tracker-mcp.toml` next to it, or by
the file `TRACKER_CONFIG` points to. `examples/tracker-mcp.example.toml` documents
every key. In short:

```toml
sheet = "Applications"

[columns]                 # field = header text (exact, else the first header starting with it)
company = "Company"       # company, role and status are required
role = "Role"
status = "Status"
applied_on = "Applied on"
followup_on = "Follow-up date"
followup_log = "Follow-up log"
salary = "Salary"         # any other key is a custom field

[statuses]
not_applied = "Not applied"
applied = "Applied"
closed = ["Rejected", "Withdrawn", "No response", "Offer"]

[followups]
first_after_days = 10
next_after_days = 14
weekend_to_monday = true
```

A missing optional column only switches off what needs it: with no follow-up
column, `followups_due` says so instead of failing elsewhere. The config is
validated strictly, so a typo in a key is an error, not a silent default.
`tracker-mcp --check` prints how the config maps the workbook, and the server
refuses to start on a bad config, with the reason in the client's log.

My own tracker is in Russian; its config is `tests/fixtures/ru.tracker-mcp.toml`.
The tests run the same code against it and against an English layout with a
different sheet, header row and set of columns.

## Known limits

- The version covers the whole file, so any edit elsewhere invalidates a pending
  write. That's fine for one person; a shared tracker would need row-level versions.
- A small window remains between the last version check and the rename. An xlsx
  file has no real locking.
- openpyxl drops cached formula values. The server sets `fullCalcOnLoad`, so Excel
  recalculates when the file is opened.
- Undo goes back one step only.
- Without a dialog, the token proves that a preview was made and nothing drifted
  since. It cannot prove the model showed that preview to the user.

## Run

Requires [uv](https://docs.astral.sh/uv/).

```sh
uv run pytest                                          # tests on a synthetic workbook
TRACKER_PATH=~/path/tracker.xlsx uv run mcp dev src/tracker_mcp/server.py   # MCP Inspector
```

Claude Code runs the server with the Terminal's permissions, which is all it needs:

```sh
claude mcp add tracker -e TRACKER_PATH=/absolute/path/to/tracker.xlsx -- /bin/sh /absolute/path/to/tracker-mcp/run.sh
```

`run.sh` exists because GUI apps on macOS don't get the shell `PATH`, so a bare
`uv` would not be found. Backups and `audit.jsonl` go to `backups-tracker-mcp/`
next to the tracker; `TRACKER_BACKUP_DIR` overrides that location.

### Claude Desktop on macOS

Claude Desktop starts a local server as its own process, without Claude's privacy
permissions. A bare `/bin/sh` then cannot read `~/Documents`, where my tracker
lives: macOS answers "Operation not permitted" and never shows a prompt. Full Disk
Access for Claude does not change that.

The fix is a small launcher app, `~/Applications/Tracker MCP.app` (`macos/launcher.c`,
about 60 lines of C, signed locally). It starts the server as a child process. A
child inherits the process macOS holds responsible, so macOS asks once whether
"Tracker MCP" may access Documents. Only this app gets that permission, not Claude
as a whole. The launcher passes stdin/stdout straight through and forwards
termination signals. It also logs to stderr whether it runs in a sandbox and
whether it can read the tracker.

`sh install.sh /absolute/path/to/tracker.xlsx` does the whole setup: it syncs
dependencies, runs the tests, and builds and signs the launcher. Then it backs up
Claude Desktop's config and adds the server to it. Quit Claude first, because
Claude rewrites its config when it quits, and the script refuses to run while
Claude is open. The resulting entry:

```json
{
  "mcpServers": {
    "tracker": {
      "command": "/Users/you/Applications/Tracker MCP.app/Contents/MacOS/tracker-mcp",
      "args": ["/bin/sh", "/absolute/path/to/tracker-mcp/run.sh"],
      "env": { "TRACKER_PATH": "/absolute/path/to/tracker.xlsx" }
    }
  }
}
```

## Layout

- `src/tracker_mcp/tracker.py` holds all file logic and knows nothing about MCP.
- `src/tracker_mcp/config.py` is the config schema (pydantic, strict) and loader.
- `src/tracker_mcp/server.py` holds the MCP tools: type hints become the input
  schema, and each tool makes one call into `Tracker`.
- `macos/` holds the launcher for Claude Desktop: `launcher.c` and the bundle's `Info.plist`.
- `tests/` contains the domain tests on my Russian layout, the same logic on an
  English layout (`test_config.py`), and a test that drives the calls through an
  in-process MCP client.
- `examples/tracker-mcp.example.toml` is a commented config to start from.

Built with the official [MCP Python SDK](https://github.com/modelcontextprotocol/python-sdk) (v2).
MIT licensed.
