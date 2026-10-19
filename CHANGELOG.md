# Changelog

## 0.4.0 — 2026-10-19

What a hand-kept tracker cannot do by itself: find its own drift, remember when each
reply came, and read the mailbox for you.

Added
- `check_tracker` and `fix_tracker`: dates kept as text or bare numbers, rows without a
  status, statuses spelled unlike the dropdown, follow-ups on a weekend or on closed
  applications, waiting rows with no follow-up date, the applied column disagreeing with
  the status, the furthest stage behind the status, replies without a reply date,
  duplicates. The fixable ones are fixed in one confirmed write, one backup, one undo.
- A reply-date column (`replied_on`; "Replied on" in the `new` template, found by `init`
  as "Reply date", "Дата ответа" and the like). `set_status` fills it on the first reply,
  with `on=` for the day the reply came; `check_tracker` proposes past dates from the
  Response column.
- `reply_times`: days from submitting to the first reply, in total and per channel,
  source or month, and how many silent applications still got a reply after 7 to 60 days.
- `match_emails` and the **Update the tracker from my e-mail** prompt: the client's mail
  tool finds the messages, the server names the application and the change (a
  confirmation, a rejection, an interview, an offer) with the e-mail's date.
- `export_followups_calendar` and `tracker-mcp calendar`: the open follow-ups as an
  `.ics` file to import into a calendar.
- `[rules] same_company_days` (28): adding or submitting to a company that already has
  an application within that many days shows a warning in the preview.

Changed
- A waiting application is silent from its last touch: the submit date or the newest
  dated follow-up log entry, whichever is later (`[rules] silence_from = "applied"` for
  the old rule). `stale_applications` shows `last_touch` and `days_silent`.
- The weekly review prompt starts with `check_tracker` and includes `reply_times`.

Fixed
- On Windows, an undo right after a write could take the same backup name as that
  write's backup, overwrite it and undo nothing: the clock there can repeat a
  microsecond. Backup names are now always new.

## 0.3.1 — 2026-10-19

Google Sheets without a Google Cloud project of your own.

Added
- `uvx tracker-mcp auth google` signs in with tracker-mcp's own Google client: Google's
  sign-in, then its file picker, and the server can open only the spreadsheets you pick
  (scope `drive.file`). No Cloud project, API keys or consent screen to set up. The client
  is built into the PyPI packages; a copy run from the source code still takes
  `--client-secret`.
- A project site with the privacy policy: https://butyavk.github.io/tracker-mcp/

Changed
- README: why a server rather than Claude editing the file directly, and the Claude
  Desktop extension as the first way to install (it reads a tracker in Documents
  without the launcher).

Fixed
- 0.3.0 said a client id alone was enough for `auth google`. Google's token endpoint
  requires a Desktop client's secret even with PKCE, so a client JSON needs it.

## 0.3.0 — 2026-10-19

The first public release: the server now works with anyone's tracker, not just mine.

Added
- Trackers in **Google Sheets** (your Google account via `tracker-mcp auth google`,
  or a service account) and in **CSV** files, besides Excel.
- On first start the server writes the config from the tracker's headers and
  dropdowns, so it runs with no setup; `describe_tracker` says what to check.
- `tracker-mcp init` (the same guess, from the command line), `new` (an empty
  tracker), `example` (every config key explained), `--tracker`, `--config`,
  `--backup-dir`, `--version`.
- `funnel`: submitted → replied → screening → interview → offer, by channel, source,
  month or any column; an optional `reached` column keeps the furthest stage, and
  `[funnel.by.<column>]` groups one column's free text. `init` flags a tracker
  without a furthest-stage column.
- `reschedule_followup`: move the next follow-up, or plan none, when nothing was sent
  (no contact yet, waiting for message credits), with the reason in the log.
- The `new` template: Location, Salary, Priority and Next step columns; "Found via"
  and "Applied via" instead of the easily confused Source and Channel.
- Prompts (weekly review, draft a follow-up) and resources (`tracker://layout`,
  `tracker://application/{row}`); titles on the server and every tool.
- `date_formats` for dates kept as text (US and UK formats are detected by `init`).
- Config edits apply on the next call, without restarting the client.
- A Claude Desktop extension (`.mcpb`), a Claude plugin with a `weekly-review` skill
  (`/plugin install tracker-mcp --marketplace butyavk/tracker-mcp`), a `server.json`
  for the MCP Registry, and evals that check a model uses the tools right and ignores
  instructions in cells.
- `tracker-mcp auth google --client-secret <client.json> --pick` signs in and opens
  Google's file picker: the server can open only the sheets you pick (scope
  `drive.file`).

Changed
- Every write needs the user's yes, enforced by the server: a confirmation dialog
  when the client has one, otherwise the token from a dry run.
- Writes are refused when openpyxl would drop parts of the workbook on save.
- Text written by the server never becomes a formula, in any format.
- On Windows, a file held by Excel, OneDrive or an antivirus gives a clear error.
- Only .xlsx and .csv files are accepted: an .xlsm would lose its macros on save, and
  .xls or .ods files get a message to save as .xlsx. A damaged file, a CSV saved as
  Windows-1252, no network, or odd Google tab names give readable errors.

## 0.2.0 and 0.1.0

Private versions for one tracker: read tools, safe writes with backups, undo and an
audit log, and the macOS launcher for Claude Desktop.
