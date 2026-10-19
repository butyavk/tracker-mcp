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

### Why a server? Can't Claude just do this?

It can, if you set it up. Claude reads a spreadsheet you attach, and with access to your
files (Claude Code, or Claude working in a folder on your computer) or the Google Drive
connector it can edit the tracker by writing a little code. Give it good instructions
(where the tracker is, what each column means, how follow-ups are counted, "back up and
ask me before every write") and it works: tracker-mcp grew out of exactly that setup. The
difference is where the rules live: in instructions the model reads and follows, or in a
program that applies them.

| | Claude with your instructions | tracker-mcp |
|---|---|---|
| Knows your tracker | From the instructions you wrote, or worked out again in each chat | Learned once from your headers and dropdowns, kept in a config |
| Follow-up dates, silent applications, funnel | The model writes code for it each time | Fixed rules in tested code: the same answer every time |
| Asks before writing | When the instructions say so and the model follows them | The server refuses any write you have not confirmed |
| Backup, undo, change log | When the model makes them | Every write |
| Dropdowns and formulas on new rows | Depends on the code written for that change | Extended to new rows, every time |
| Excel has the file open, or it changed meanwhile | Checked if the model thinks of it | The write is refused |
| Text pasted from job posts and e-mails | Text starting with "=" becomes a formula unless the code prevents it | Always written as text |
| What earlier edits left behind (text dates, rows without a status) | Found if you ask, fixed by code written for it | `check_tracker` finds it, `fix_tracker` fixes it in one confirmed write |
| When each reply came | Known only if someone wrote it down | Recorded on every reply; `reply_times` shows when replies stop coming |
| Replies in your mailbox | The model reads them and decides which row each is about | `match_emails` names the row and the change, the same way every time |
| Each request | Write, run and check a script | One tool call |
| Free-form analysis, charts, reshaping the sheet | Anything | Not its job: Claude still does that alongside |
| Where | Where Claude can run code on your files, or on a copy you upload | Any MCP client: Claude Desktop (in a normal chat), Claude Code, Cursor, ...; not in a browser or on a phone |

So tracker-mcp mostly does not give Claude new abilities. It gives the everyday part
of a job search (record what happened, see what is due, see what works) fixed rules and
fixed safety, with no instructions to write. The one thing Claude cannot do later is
know when a reply came if nobody wrote it down; the server writes it down as it happens.

- **Rules belong in code, not in a prompt.** "Which follow-ups are due?" is date
  arithmetic with your rule: 10 days after applying, then every 14, weekends moved to
  Monday, nothing for closed applications. A model gets it right most of the time, but it
  works it out anew in each chat, and the slips are the quiet kind: a follow-up for a
  rejected application, a date on a Sunday, a status spelled differently from your
  dropdown. Here the rule is code covered by tests. The same goes for silent applications
  and the funnel.
- **Safety the model cannot skip.** You see the exact cells that will change, and your
  yes is checked by the server, not taken on the model's word. That matters most where no
  one re-reads the instructions: a scheduled weekly review, a long chat, a note pasted from
  a job post that tells the model to do something.
- **Fewer steps.** One tool call instead of writing and running a script, so answers come
  quicker and the model's attention stays on your question.
- **The same tracker everywhere.** Every chat, a scheduled review and every client you
  use work on the same file, and each change lands in it right away.

Your spreadsheet stays yours: you keep opening it in Excel or Sheets, sorting and
editing it. Nothing about a particular tracker is in the code; the server reads your
headers and dropdowns and keeps what it learned in a small config file.

You don't need tracker-mcp if you only want a one-time analysis of your applications,
if you rarely update the tracker, or if you only use Claude in a browser or on your phone
(a server like this runs on your computer). If Claude already keeps your tracker well with
instructions you wrote, what you gain is reliability, not new features.

## Quick start

You need [uv](https://docs.astral.sh/uv/getting-started/installation/), which also
fetches Python for you. Pick your client:

**Claude Desktop, the easy way:** download `tracker-mcp-<version>.mcpb` from the
[latest release](https://github.com/butyavk/tracker-mcp/releases), double-click it,
press **Install**, choose your tracker file in *Your tracker*, **Save**, and switch the
extension on. On a Mac this also avoids the [Documents folder problem](#macos-claude-desktop-and-your-documents-folder).

**Claude Desktop, by hand.** Open Settings → Developer → Edit Config and add:

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

**Claude Code:** as a plugin (the server plus a `weekly-review` skill; it asks for the
path to your tracker):

```
/plugin install tracker-mcp --marketplace butyavk/tracker-mcp
```

or just the server:

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
| **Google Sheets** | `TRACKER_PATH=https://docs.google.com/spreadsheets/d/<id>/edit` | Sign in once and pick the sheet, below |

### Google Sheets

Sign in once and pick your tracker; there is nothing to set up in Google Cloud:

```sh
uvx tracker-mcp auth google
```

Your browser opens Google's sign-in and then Google's file picker. Choose the spreadsheet
with your applications. tracker-mcp writes its config for that sheet and prints what to
add to your client: the same setup as for a file, with the sheet's URL as `TRACKER_PATH`.

- **What tracker-mcp can reach:** only the files you pick in Google's picker (or that it
  creates), through Google's `drive.file` permission. It cannot see your other sheets or
  files, your e-mail or your contacts. To use another sheet later, run the command again
  and pick it.
- **Where the sign-in is kept:** on your computer, in `~/.tracker-mcp/google-token.json`
  (readable by you only), with backups of changed cells in `~/.tracker-mcp/backups/`.
  There is no tracker-mcp server; the program talks to Google directly. To withdraw
  access, remove tracker-mcp at [myaccount.google.com/permissions](https://myaccount.google.com/permissions).
  The [privacy policy](https://butyavk.github.io/tracker-mcp/privacy.html) has the details.
- **On Google's screens** you sign in "to continue to tracker-mcp", with links to its
  [home page](https://butyavk.github.io/tracker-mcp/) and privacy policy. Google has not
  reviewed the app's brand, so it shows no logo; the access it asks for (`drive.file`) is
  one Google lets apps request without a review.
- **No browser on that computer** (a server over SSH): `uvx tracker-mcp auth google --no-browser`
  prints the link; the browser must still reach `127.0.0.1` on that computer to finish,
  so forward the port or sign in where the browser is and copy `~/.tracker-mcp/`.

**Your own Google client instead** (a fork, a company account that blocks outside apps,
or if you prefer not to rely on this project's client): in the
[Google Cloud console](https://console.cloud.google.com/), create a project, enable the
**Google Sheets API** (and the **Google Picker API** for `--pick`), configure the consent
screen (External), create an OAuth client of type **Desktop app** and download its JSON:

```sh
uvx tracker-mcp auth google --client-secret ~/Downloads/client_secret_….json --pick   # only the sheets you pick
uvx tracker-mcp auth google --client-secret ~/Downloads/client_secret_….json          # all your spreadsheets
```

Without `--pick` the permission (`spreadsheets`) covers every spreadsheet you can open,
although the server opens only the one in `TRACKER_PATH`; then run
`uvx tracker-mcp init "https://docs.google.com/spreadsheets/d/<id>/edit"` to write the
config. While your consent screen is in *Testing*, Google ends the sign-in after 7 days;
publish it (Audience → Publish app) to keep it.

**Or a service account:** create one, download its key, share your sheet with the
service account's e-mail, and set `GOOGLE_APPLICATION_CREDENTIALS=/path/key.json` in the
server's environment. The server can then reach only the sheets you share with it.

## What Claude can do with it

| Tool | Kind | What it does |
|---|---|---|
| `describe_tracker` | read | How the tracker is laid out: fields → columns, allowed and closed statuses, follow-up rule |
| `list_applications` | read | Filter by status, company/role text, submit dates |
| `get_application` | read | Every column of one application (by text or row) |
| `followups_due` | read | Follow-ups due within N days, overdue first; submitted rows with no follow-up date |
| `stale_applications` | read | Still waiting, silent N+ days since the last touch (the submit date or the newest dated follow-up log entry), no follow-up planned |
| `summary` | read | Counts by status and month; a cohort for a date range |
| `funnel` | read | Submitted → replied → screening → interview → offer, in total and by channel, source, month or any column |
| `check_tracker` | read | What hand edits left behind: dates kept as text, rows without a status, statuses not in the dropdown, weekend follow-ups, closed rows with a planned follow-up, the furthest stage behind the status, replies without a date, duplicates |
| `reply_times` | read | Days from submitting to the first reply (median, 75th percentile), in total and by group, and how many silent applications still got a reply after 7, 14, 21, 30, 45, 60 days |
| `match_emails` | read | Which application each e-mail is about (confirmation, rejection, interview, offer) and the write that would record it |
| `recent_writes` | read | Audit log of the server's writes |
| `add_application` | write | New row with the not-applied status; other columns via `extra`; refuses duplicates |
| `mark_applied` | write | Submitted: date, channel, status, follow-up date by your rule |
| `set_status` | write | New status from your dropdown; a closed status cancels the pending follow-up; the first reply fills the reply date (`on`, default today) |
| `log_followup` | write | Log a sent follow-up, schedule the next one or close |
| `reschedule_followup` | write | Move the next follow-up, or plan none, when nothing was sent (no contact yet, waiting for credits) |
| `edit_application` | write | Correct fields of one row: a wrong or missing date, a typo, the furthest stage, a custom column; dates are checked, the reason goes to the notes |
| `add_column` | write | Add a column after the last one, styled like its neighbour, and map it in the config (`replied_on`, `reached`, your own field); undo removes both |
| `close_stale` | write | Close stale applications (or chosen rows) in one write, one backup, one undo step |
| `fix_tracker` | write | Fix what `check_tracker` can fix, in one write, one backup, one undo step |
| `export_followups_calendar` | file | The open follow-ups as an `.ics` file next to the tracker (the tracker is not touched) |
| `undo_last_write` | write | Restore the state before the last write, if nothing changed since |

Three prompts, for clients that show MCP prompts: **Weekly review** checks the
tracker and goes through follow-ups, stale applications, reply times, the funnel and
this week's submissions, and ends with a plan; **Update the tracker from my e-mail**
has the client's own mail tool find replies from companies, matches them with
`match_emails` and proposes the changes; **Draft a follow-up** writes a short message
for one row. Two resources: `tracker://layout` and `tracker://application/{row}`.

`add_application` and `mark_applied` warn when the same company already has an
application submitted within 28 days (`[rules] same_company_days`); the warning shows
in the preview and does not stop the write.

### Reply dates and when to stop waiting

Give the tracker a reply-date column (`Replied on`, `Дата ответа`; the `new` template
has one) and `set_status` fills it on the first reply: a rejection or any funnel
stage, with `on=` for the e-mail's date. `check_tracker` fills past ones from the
Response column (the earliest date there that is not about the application being
received) for you to approve. Then `reply_times` says how long your replies take per
channel, and after how many silent days they stopped coming: the point where closing
an application costs nothing. Without the column, it uses only the replies recorded
through the server.

### Follow-ups in your calendar

`tracker-mcp calendar --tracker ~/Documents/job-tracker.xlsx` (or the
`export_followups_calendar` tool) writes the open follow-ups as an `.ics` file next to
the tracker: one all-day event each, overdue ones on today, with the contact and the
last log line. Open the file to import it into Calendar, Outlook or Google Calendar.
It is a snapshot: after the dates change, export and import again (each event keeps
its ID; importing into a calendar of its own makes an old copy easy to clear). If your
client has a calendar tool, it can put the follow-ups there directly.

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

**The easy way: use the Claude Desktop extension** (`tracker-mcp-<version>.mcpb` from
the [releases](https://github.com/butyavk/tracker-mcp/releases)). Claude Desktop runs an
extension under its own permissions, so a tracker in Documents works: in our test
(Claude Desktop 2.42304, October 2026) the extension read a tracker in `~/Documents`
with nothing else to set up. If macOS asks whether Claude may access the folder, allow it.

The rest of this section is about adding the server to Claude Desktop's config by hand.
Claude Desktop starts such servers without its own privacy permissions. If your
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
