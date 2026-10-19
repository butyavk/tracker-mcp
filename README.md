# tracker-mcp

<!-- mcp-name: io.github.butyavk/tracker-mcp -->

An MCP server for the spreadsheet where you track your job applications: Excel,
CSV or Google Sheets. Connect it to Claude (or any MCP client), and the assistant
keeps the tracker for you: it answers from it, reminds you what is due, and records
what happened, with your yes for every change.

> **You:** Which follow-ups are due this week?
>
> **Claude:** Three. Acme (applied 17 Sep): nudge Sarah Kim on LinkedIn, overdue by
> two days. Globex: due Thursday, no contact yet. Initech: due Friday.
>
> **You:** I can't reach anyone at Globex until Monday. Move it.
>
> **Claude:** Row 8 · Globex — Head of Engineering. Follow-up date: 23 Oct → 27 Oct;
> follow-up log: + "follow-up moved to 2026-10-27: no contact yet". Write it?
>
> **You:** Yes.
>
> **Claude:** Done. A backup was taken first; say "undo" to revert.

Also: "Which channel gets me the most replies?", "Initech rejected me", "I applied to
Hooli via a referral today", "close everything that hasn't answered in 45 days".

### Why not just give Claude the spreadsheet?

Claude can already read a spreadsheet you attach, and its Google Drive connector can
open your Sheets. tracker-mcp adds what a job search needs on top:

- **Rules, not guesses.** Follow-up dates by your rule (10 days after applying, then
  every 14, weekends moved to Monday), applications that went silent, and a funnel
  (submitted → replied → screening → interview → offer) by channel, source or month
  that remembers the furthest stage an application reached.
- **Safe writes.** You see the exact cells that will change, and the server, not the
  model, checks your yes. A backup before every write, an audit log, undo; nothing is
  written while Excel has the file open or after someone else changed it.
- **Always there.** The same tracker in every chat and every session, with no upload,
  so several conversations (or a scheduled weekly review) can keep one tracker up to date.

Your spreadsheet stays yours: you keep opening it in Excel or Sheets, sorting and
editing it. Nothing about a particular tracker is in the code; the server reads your
headers and dropdowns and keeps what it learned in a small config file.

## Quick start

You need [uv](https://docs.astral.sh/uv/getting-started/installation/), which also
fetches Python for you. Pick your client:

**Claude Desktop.** Open Settings → Developer → Edit Config and add:

```json
{
  "mcpServers": {
    "tracker": {
      "command": "uvx",
      "args": ["tracker-mcp"],
      "env": { "TRACKER_PATH": "/Users/you/Documents/job-tracker.xlsx" }
    }
  }
}
```

On Windows the path looks like `C:\\Users\\you\\Documents\\job-tracker.xlsx`
(double backslashes in JSON). Restart Claude Desktop. If Claude can't find `uvx`,
put its full path in `command` (`which uvx` / `where uvx`). On macOS, if your
tracker is in Documents, Desktop or Downloads, see [macOS privacy](#macos-claude-desktop-and-your-documents-folder).

A one-click Claude Desktop extension (`tracker-mcp-<version>.mcpb`) is attached to each
[release](https://github.com/butyavk/tracker-mcp/releases): double-click it and pick your file.

**Claude Code:**

```sh
claude mcp add tracker -e TRACKER_PATH="$HOME/Documents/job-tracker.xlsx" -- uvx tracker-mcp
```

**Cursor, VS Code, others:** a stdio server with command `uvx`, arguments
`["tracker-mcp"]` and the environment variable `TRACKER_PATH`. Cursor reads the same
JSON as Claude Desktop from `~/.cursor/mcp.json`.

**From a clone** (to change the code, or before a release is out):

```sh
git clone https://github.com/butyavk/tracker-mcp && cd tracker-mcp && uv sync
```

and use `"command": "uv", "args": ["--directory", "/path/to/tracker-mcp", "run", "tracker-mcp"]`.

### First start

On its first start the server reads your tracker, finds the sheet and the header
row, maps headers to fields (English and Russian headers are understood), works out
what each status in your dropdown means, and writes `job-tracker.tracker-mcp.toml`
next to the tracker. It does not touch the tracker itself. Ask Claude to describe
your tracker: it shows the mapping and the guesses worth checking. Edit the config
whenever you like; changes apply on the next request, no restart needed.

To see the guess before anything is written, run it yourself:

```sh
uvx tracker-mcp init ~/Documents/job-tracker.xlsx     # writes the config, lists what to check
uvx tracker-mcp --check --tracker ~/Documents/job-tracker.xlsx   # how the config maps the tracker
```

No tracker yet? `uvx tracker-mcp new ~/Documents/job-tracker.xlsx` creates an empty
one with dropdowns, a filter, a summary sheet and its config.

### A weekly routine

The server answers when asked; it does not send reminders. Two habits make it pay off:

- **Once a week**, run the **Weekly review** prompt (or just ask "review my job search
  for this week"): what is due, what went silent, how the funnel moves, and a plan. If
  your client can run scheduled tasks, schedule that sentence for Monday morning.
- **After each event**, tell Claude in a sentence: "applied to X via Y", "X rejected
  me", "sent the follow-up to X", "can't follow up with X until Friday". It finds the
  row, shows the change and writes it when you say yes.

## Where your tracker can live

| | How to point to it | Notes |
|---|---|---|
| **Excel** (.xlsx) | `TRACKER_PATH=/path/tracker.xlsx` | Dropdowns, filters and summary formulas are kept and extended to new rows |
| **CSV** | `TRACKER_PATH=/path/tracker.csv` | Exports from Notion, Huntr, Teal or Sheets. The delimiter, the encoding (UTF-8 or Windows-1252) and line endings are kept; statuses come from the config |
| **Google Sheets** | `TRACKER_PATH=https://docs.google.com/spreadsheets/d/<id>/edit` | Needs a one-time sign-in, below |

### Google Sheets

The server talks to the Sheets API directly, under your Google account or a
service account. Either way, Google needs a project of yours to know who is asking:

1. In the [Google Cloud console](https://console.cloud.google.com/), create a project
   (any name) and enable the **Google Sheets API** for it.
2. **Your own account:** under *Google Auth Platform*, configure the consent screen
   (External, add yourself as a test user), then create an OAuth client of type
   **Desktop app** and download its JSON. Sign in once:

   ```sh
   uvx tracker-mcp auth google --client-secret ~/Downloads/client_secret_….json
   ```

   The sign-in is kept in `~/.tracker-mcp/google-token.json`. While the consent screen
   is in *Testing*, Google ends the sign-in after 7 days; switch it to *In production*
   (Google then warns that the app is unverified, which is expected for your own
   client) or sign in again.

   **Or a service account:** create one, download its key, share your sheet with the
   service account's e-mail, and set `GOOGLE_APPLICATION_CREDENTIALS=/path/key.json`
   in the server's environment. This way the server can reach only the sheets you
   share with it.
3. `uvx tracker-mcp init "https://docs.google.com/spreadsheets/d/<id>/edit"` writes the
   config to `~/.tracker-mcp/<id>.tracker-mcp.toml`. Then use the URL as `TRACKER_PATH`.

Signing in with your own account asks for access to your spreadsheets
(`spreadsheets` scope). The server opens only the sheet in `TRACKER_PATH`, but the
permission covers all of them; if that matters, use a service account.

## What Claude can do with it

| Tool | Kind | What it does |
|---|---|---|
| `describe_tracker` | read | How the tracker is laid out: fields → columns, allowed and closed statuses, follow-up rule |
| `list_applications` | read | Filter by status, company/role text, submit dates |
| `get_application` | read | Every column of one application (by text or row) |
| `followups_due` | read | Follow-ups due within N days, overdue first; submitted rows with no follow-up date |
| `stale_applications` | read | Still waiting, submitted N+ days ago, no follow-up planned |
| `summary` | read | Counts by status and month; a cohort for a date range |
| `funnel` | read | Submitted → replied → screening → interview → offer, in total and by channel, source, month or any column |
| `recent_writes` | read | Audit log of the server's writes |
| `add_application` | write | New row with the not-applied status; other columns via `extra`; refuses duplicates |
| `mark_applied` | write | Submitted: date, channel, status, follow-up date by your rule |
| `set_status` | write | New status from your dropdown; a closed status cancels the pending follow-up |
| `log_followup` | write | Log a sent follow-up, schedule the next one or close |
| `reschedule_followup` | write | Move the next follow-up, or plan none, when nothing was sent (no contact yet, waiting for credits) |
| `close_stale` | write | Close stale applications (or chosen rows) in one write, one backup, one undo step |
| `undo_last_write` | write | Restore the state before the last write, if nothing changed since |

Two prompts, for clients that show MCP prompts: **Weekly review** goes through
follow-ups, stale applications, the funnel and this week's submissions, and ends
with a plan; **Draft a follow-up** writes a short message for one row. Two
resources: `tracker://layout` and `tracker://application/{row}`.

Read tools are marked read-only. Writes are marked destructive where they
overwrite cells, and all of them take `dry_run`, `expected_version` and `confirm`.

## How a write stays safe

1. **The server checks your yes.** Before a write it computes the plan, the exact
   cell changes at the current version of the tracker. A client with confirmation
   dialogs (MCP elicitation) shows it to you there; otherwise the write needs the
   token of a dry run that you saw, which fingerprints the changes and the version.
   The write lands exactly as approved, or not at all.
2. **Nobody's edit is lost.** Each read returns a `version` (a hash of the file, or of
   the sheet's values); a write against an old version is refused, and so is a write
   while Excel has the file open.
3. **Write, check, keep a backup.** Files are written to a temporary copy, re-read
   cell by cell, backed up and swapped in atomically; a Google Sheet gets one update
   of the changed cells, is read back and rolled back on any difference.
4. **Audit and undo**: one line per write; undo only if nothing changed since.
5. **Your data is data**: text never turns into a formula, and the model is told not
   to follow instructions found in cells (a pasted job post). See [SECURITY.md](SECURITY.md).

Allowed statuses come from the status column's dropdown. Columns are found by
header text, not by letter, so you can reorder them. When a row is added to an
.xlsx file, the server stretches dropdowns, the filter and summary formulas to it.

## Config

`job-tracker.xlsx` is described by `job-tracker.tracker-mcp.toml` next to it (for a
Google Sheet, `~/.tracker-mcp/<id>.tracker-mcp.toml`), or by the file `TRACKER_CONFIG`
points to. `uvx tracker-mcp example` prints a config with every key explained
([examples/tracker-mcp.example.toml](examples/tracker-mcp.example.toml)). In short:

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
rejected = ["Rejected"]

[statuses.stages]         # funnel stages after a reply, in order
screening = ["Screening"]
interview = ["Interview", "Final"]
offer = ["Offer"]

[followups]
first_after_days = 10
next_after_days = 14
weekend_to_monday = true
```

A missing optional column only switches off what needs it: with no follow-up
column, `followups_due` says so. The config is validated strictly, so a typo in a
key is an error, not a silent default. If an edit breaks the config, every request
says why until it is fixed.

The funnel counts a rejection as a reply and needs to know your stages
(`[statuses.stages]`). The status column only holds the current state, so add a
`reached` column ("Furthest stage") for the furthest stage, or a rejection after an
interview counts as just a reply; `set_status` moves it forward, never back.
Free-text columns such as the source are grouped by their text up to the first
bracket or dash ("Ashby (direct form)" → "Ashby"), or by substrings you list in
`[funnel.groups]` (all columns) or `[funnel.by.source]` (one column).

| Variable | Option | |
|---|---|---|
| `TRACKER_PATH` | `--tracker` | The tracker: a file path or a Google Sheets URL (required) |
| `TRACKER_CONFIG` | `--config` | The config, if not next to the tracker |
| `TRACKER_BACKUP_DIR` | `--backup-dir` | Backups and `audit.jsonl` (default: `backups-tracker-mcp/` next to the tracker; `~/.tracker-mcp/backups/<id>` for a Google Sheet) |
| `GOOGLE_APPLICATION_CREDENTIALS` | | A service account key, for Google Sheets |
| `TRACKER_MCP_HOME` | | Where the Google sign-in and Sheets configs live (default `~/.tracker-mcp`) |

## Known limits

- The version covers the whole tracker, so any edit elsewhere in it invalidates a
  pending write. Fine for one person; a shared tracker would need row-level versions.
- A small window remains between the last version check and the write: neither an
  .xlsx file nor the Sheets API offers real locking.
- openpyxl cannot keep some parts of newer Excel files, such as dropdowns that take
  their list from another sheet and are stored in an extension. If loading the file
  warns that a part would be removed, the server refuses writes to that workbook
  rather than drop it, and says so. Reading works.
- openpyxl drops cached formula values; the server asks Excel to recalculate on open.
- In a CSV file, text the server writes that starts like a formula (`=`, `+`, `@`,
  `-1…`) gets a leading apostrophe, so Excel shows it as text.
- The config is named after the tracker without its extension, so `jobs.xlsx` and
  `jobs.csv` in one folder would share one; give one of them `TRACKER_CONFIG`.
- Only .xlsx and .csv files: an .xlsm would lose its macros on save, and .xls, .ods or
  Numbers files need saving as .xlsx first.
- `init` understands English and Russian headers and status words. In other
  languages it maps less and lists the rest under "check by hand".
- Undo goes back one step only. For a Google Sheet it restores values, not formatting.
- Without a dialog, the token proves that a preview was made and nothing drifted
  since. It cannot prove the model showed you that preview.

## macOS: Claude Desktop and your Documents folder

Claude Desktop starts local servers without its own privacy permissions. If your
tracker is in Documents, Desktop or Downloads, the server may fail with "Operation
not permitted" in its log (Settings → Developer), and
macOS never shows a prompt. Full Disk Access for Claude does not change that.
Moving the tracker to another folder (say `~/JobSearch`) is the simplest fix.

To keep it where it is, `sh install.sh /absolute/path/to/tracker.xlsx` (from a clone)
builds a tiny launcher app, `~/Applications/Tracker MCP.app` (`macos/launcher.c`,
about 60 lines of C, signed locally). It starts the server as its child, so macOS
asks once whether "Tracker MCP" may access Documents, and only this app gets that
permission. The script also runs the tests, backs up Claude Desktop's config and
adds the server to it; quit Claude first.

## Development

```sh
uv sync
uv run pytest                     # the tests use synthetic trackers and a fake Sheets API
uvx ruff check src tests evals
npx @modelcontextprotocol/inspector uv --directory . run tracker-mcp --tracker /path/tracker.xlsx
ANTHROPIC_API_KEY=… uv run python evals/run_evals.py   # does a model use the tools right? (a few cents)
```

- `src/tracker_mcp/tracker.py` is the tracker logic; it knows nothing about MCP or storage.
- `stores.py` reads and writes .xlsx and .csv files; `gsheets.py` does Google Sheets.
- `config.py` is the config schema (pydantic, strict); `scaffold.py` is `init` and `new`.
- `server.py` holds the MCP tools, prompts, resources and the confirmation flow.
- `evals/` checks that a model picks the right tools, asks before writing and
  ignores instructions planted in a cell.
- `macos/` is the launcher for Claude Desktop; `mcpb/` builds the Claude Desktop extension.

See [CONTRIBUTING.md](CONTRIBUTING.md) for how releases are made. Built with the
official [MCP Python SDK](https://github.com/modelcontextprotocol/python-sdk). MIT licensed.
