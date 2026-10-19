# Changelog

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
- A Claude Desktop extension (`.mcpb`), a `server.json` for the MCP Registry, and
  evals that check a model uses the tools right and ignores instructions in cells.

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
