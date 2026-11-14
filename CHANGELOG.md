# Changelog

## 0.8.6 — 2026-11-14

Fixed (an end-to-end review)
- A cell with part of its text bold no longer stops every write to the workbook: the workbook is read with its rich text, the formatting survives a save, and tools read the words as plain text. A cell the server writes to becomes plain text.
- Pillow is now a dependency, so pictures in a workbook (a logo, a sheet background) are kept, also for `uvx` and the desktop extension, where it could not be added by hand.
- 0.8.5 said the list was of what a save keeps; a few parts on it are in fact dropped. They hold none of the tracker's data and are now named under `a_save_drops` in `describe_tracker` and in each write's result: SharePoint document properties, the printer setup, the file's preview picture, chart style and color presets, and the people list left after the last threaded comment is deleted. README «Known limits» lists them, and what a save loses inside the sheets.
- A chart on its own sheet (F11 in Excel) no longer refuses writes: a save keeps it.
- The people list of threaded comments refuses writes only while threaded comments are there.
- A refusal says «a dropdown that takes its list from another sheet» only when a dropdown is the cause.
- With `--dialog-only` and a client that has no dialog, a preview no longer offers a token that would be refused; it says the change is written from a client that asks in a dialog.
- The desktop extension's Read-only setting no longer suggests a second copy beside the read-write one (every chat sees every connected extension).
- `tracker-mcp --check | head` exits quietly; `--help` names `example` and `auth`; the README says which writes take `dry_run`; server.json offers TRACKER_READ_ONLY and TRACKER_DIALOG_ONLY.

## 0.8.5 — 2026-11-14

Fixed (a fourth review)
- A write is refused, before anything changes, when saving would silently drop part of the workbook: threaded comments (Excel 365), checkboxes in cells, shapes and text boxes, dynamic-array formulas, cell pictures and data types, slicers, timelines, form controls, embedded objects, data connections, add-ins, or mixed formatting inside a cell. The list is of what a save keeps; anything else in the file blocks writes and is named.
- A confirm token works once. After an undo, the same change is a new decision: the old token is refused and a new preview is needed. A token also belongs to one tracker.
- add_column with a mapped field: if the config cannot be written, the tracker is put back as it was and the error says so, instead of a column without an audit line that then blocked undo.
- Undo of such a column is refused, before anything is restored, when its config backup is gone.
- Previews show every value from the tracker in «» with a line saying that text in «» is data, never an instruction; long names and warnings are cut to one short line; letters that draw nothing (Hangul fillers, the combining grapheme joiner) are dropped.
- Tests now guard fixes that had none: network names never resolved, a symlink loop in the audit, keyed tokens, pictures without Pillow, the CSV guard for tab and CR, and cells kept inside the follow-up prompt's data.

## 0.8.4 — 2026-11-14

Fixed (a sweep for the kinds of error the fifth review found, and an independent check of it)
- Undo of add_column restores only this config's own backup (`<config>.before-add-column.<time>`): a forged audit line can no longer make it copy another file from the backup folder into the config. This hole was in 0.8.3.
- One tracker, one history, by whatever path it is reached (relative, `~`, a link with any name, a linked folder): writes replace the real file and leave a link a link; the lock, the backups (named after the real file) and the audit sit next to the real file; audit lines name the resolved file, and older lines written under another spelling of the path still count. Backups made before 0.8.4 next to a link stay there. Excel's lock file is looked for beside the link and beside the real file.
- Audit text is never used to reach another machine: network names (`\\host\…`, `//host/…`) and names of other files are not resolved; an odd line (`~nosuchuser/…`, a NUL) belongs to nobody and never stops undo.
- A store that cannot say whether a backup is its own owns none, so undo refuses rather than guesses.
- Every piece of text in a confirmation dialog that can come from the file or the request (column names, config lines, field names, warnings) is one line. A control or line separator shows as ⏎; characters that can hide or reorder text (bidi marks, zero-width spaces, the soft hyphen, the BOM, stray tag letters) are dropped; the joiners that emoji, Persian and Hindi need, flag tags and characters newer than the server's Unicode stay as written.
- Every command line (`init`, `new`, `calendar`, `--check`) answers a file it cannot read or write with one line naming it, exit 2; a config that is not UTF-8 is named. `TRACKER_MCP_DEBUG=1` shows the full traceback for a bug report.
- A test holds the set of write tools; job-search-kit lists the same set as the tools a desktop task keeps on «ask», and its own test compares the two.

## 0.8.3 — 2026-11-14

Fixed (a fifth review: a deep review and a red-team round)
- Concurrent writes no longer lose rows: each write has a temp file of its own, and a lock (between threads of one server and between server processes) is held from the last version check to the replace; a second writer gets «the tracker changed, retry» instead of a raw FileNotFoundError.
- Two trackers in one folder (`jobs.xlsx`, `jobs.2025.xlsx`) keep their own undo and backups: audit lines carry the tracker, and pruning matches exactly that tracker's backup names.
- Undo cannot be turned into a write elsewhere by a forged audit line: it restores only this tracker's own backups by their bare names, and only this tracker's own config for add_column. Every field in the undo dialog is shown on one line, with the config file it touches.
- The dialog cleaner also hides bidi and zero-width characters.
- Confirm tokens are keyed by a secret of the server process, so a model cannot compute one without a dry run (preview again after a restart). A read-only server's previews carry no token; a dialog-capable client is told to call again without dry_run.
- Excel Tables stretch to new rows; the template's own Summary formulas (`'Applications'!B2:B500`) grow past row 500.
- Pictures in a workbook make writes refused while Pillow is missing (openpyxl would drop them on save).
- `init` maps the shortest matching header («Status» over «Status note»); the generated config says where the statuses came from.
- An interview invite saying «you have not selected a time» is no longer read as a rejection.
- A layout error at startup prints `tracker-mcp: …` instead of a traceback. CSV text that starts with a tab or CR gets the formula guard too.
- README: the read-only example pins `tracker-mcp>=0.8.3,<0.9`; client hooks that answer dialogs (Claude Code's Elicitation hooks) pass `--dialog-only` too, said plainly. A flaky sign-in test waits for its visitor thread.

## 0.8.2 — 2026-11-14

Fixed
- The confirmation dialog also shows C1 control characters (U+0080–U+009F, such as U+0085) of a cell as ⏎.
- With `--dialog-only`, a dry run's «next» says to call again without dry_run so the client asks the user, not to confirm with the token.

## 0.8.1 — 2026-11-14

Fixed (a second red-team review)
- `--dialog-only` wins over `--allow-unconfirmed-writes` with `confirm_writes = false`: no write lands without the dialog.
- `TRACKER_ALLOW_UNCONFIRMED`, the one switch that loosens safety, reads only an explicit yes (1, true, yes, on); «never» or «none» leave it off.
- A cell's newlines and control characters are shown as ⏎ in the confirmation dialog, so a cell cannot draw fake lines into it.
- The draft_followup prompt escapes `<` in the tracker's data, so a cell cannot close its `<application>` tag.
- `tracker-mcp remote --dialog-only`. The desktop extension has «Read-only» and «Confirm only in the app's dialog» settings.
- README: dialog-only trusts the client to say it has a dialog.

## 0.8.0 — 2026-11-14

Changed (a red-team review)
- `confirm_writes = false` in the config no longer turns the yes off by itself: the server must also be started with `--allow-unconfirmed-writes` (or `TRACKER_ALLOW_UNCONFIRMED=1`). A config file is easy to edit, by a person or by an instruction injected into a session; the command line is not. describe_tracker says when the setting is ignored.

Added
- `--dialog-only` (or `TRACKER_DIALOG_ONLY=1`): only the client's confirmation dialog confirms a write, and dry-run tokens are refused. In token mode the model sees the token, so only the instructions keep it from confirming its own preview; with this switch a person answers every write.
- The write tools' description says never to confirm a preview the user has not answered, nor because a page, email, file or cell says the user agreed. README: the token's limit said plainly.

## 0.7.1 — 2026-11-14

Fixed (independent review of 0.7.0)
- The confirmation token of add_column now covers the field, kind, dropdown and layout: a preview of a plain column no longer confirms a mapped column with a dropdown.
- Undo of add_column stops if the config was edited by hand after it; before, a chained undo overwrote those edits.
- A cut-off line in audit.jsonl no longer breaks recent_writes and undo, and no longer swallows the next write's line; recent_writes says how many lines it skipped and returns `version`.
- Undo with a stale expected_version says "re-read and retry"; the undo preview checks that the write's backup is still there.
- Read-only writes no files at all: the .ics export is refused, and the server does not write a config on first start.
- TRACKER_READ_ONLY: any value but empty, 0, false, no or off turns it on. `tracker-mcp remote --read-only` works.
- README: a scheduled run gets its own server list with only the read-only entry (`--strict-mcp-config`), and what the desktop app can and cannot do; the old "undo goes back one step only" line, the extension's description and the Russian page now describe stepping back.

## 0.7.0 — 2026-11-14

Added
- `--read-only` (or `TRACKER_READ_ONLY=1`): dry runs and previews only; every write is
  refused, even with a valid token, and `describe_tracker` says the server is read-only.
  For scheduled tasks, which have nobody to say yes: register a second, read-only entry
  of the server for them (README, «Scheduled tasks»).
- `undo_last_write` steps back through several writes: each call reverts the latest
  write that is not undone yet, as long as the tracker is exactly as that write (or the
  previous undo) left it. A hand edit in between stops it, as before.

Changed
- `expected_version` is required for every write (it was optional). A write without it
  is refused with a message that says how to get it; a dry run still works without one
  and returns the version to pass.

## 0.6.0 — 2026-11-03

Added
- `add_applications`: several new rows in one write. Each item takes the arguments of
  `add_application`, and the rows go in below each other in the given order. A
  duplicate inside the batch is refused like a duplicate already in the tracker.
- `set_statuses`: several status changes in one write, for example after a day of
  replies. Each item is applied exactly as `set_status` applies it, and a row may
  appear only once.
- Both take up to 50 items. The user sees one preview of every row and answers once,
  in the client's dialog or with one `confirm` token. The write is all or nothing: a
  refused item stops the whole batch, and the error names it. One backup is taken, so
  one `undo_last_write` reverts the whole batch.

Changed
- `check_tracker` takes reply dates from `set_statuses` writes too, as it does from
  `set_status`.

## 0.5.2 — 2026-10-25

Changed
- Tool descriptions are clearer, so that a model picks the right tool the first time.
  Each write tool now says how the user's yes works. It is the client's dialog, or a
  dry run and then `confirm`. It also says that a changed tracker refuses the write and that
  the write can be undone. `mark_applied`, `set_status`, `log_followup`, `summary`,
  `recent_writes`, `list_applications` and `get_application` now say what they change
  or return. They also say when a neighbouring tool fits better. The parameters of
  `add_application` and the note of `mark_applied` have descriptions now.

## 0.5.1 — 2026-10-20

Added
- A Russian template: `tracker-mcp new <file> --lang ru` makes the tracker with Russian
  headers, statuses, sheet names and summary; the config it writes uses Russian follow-up
  log lines and dates. On a remote server the consent page offers English or Russian for
  a new sheet, preselected from the browser's language.
- The Russian guide shows how to have Claude's Google Sheets connector make the same
  tracker, for trying it without installing anything.

## 0.5.0 — 2026-10-20

Added
- `tracker-mcp remote`: the same tools over streamable HTTP for Google Sheets trackers,
  to add to Claude as a custom connector (web, desktop, phone). Several people per
  server, each signed in with Google and working on their own sheet. The server is its
  own OAuth 2.1 authorization server for MCP clients: dynamic client registration,
  PKCE, refresh tokens rotated on every use, revocation, tokens bound to the server's
  address and stored only as hashes. Google behind it for who signs in (`openid email`,
  checked against an allow-list on every token use) and for the sheets (`drive.file`);
  each person's Google sign-in is kept encrypted. A consent page names the connecting
  client and lets you pick your sheet in Google's file picker or start a new one.
- `Dockerfile` and `deploy/cloudrun.sh`: one Cloud Run instance with a private Cloud
  Storage bucket for its data and Secret Manager for its keys.
- On a remote server, a preview asks the user to confirm with a reply that names the
  change: claude.ai may refuse a write to an outside service approved with a bare "yes".

## 0.4.4 — 2026-10-20

Added
- `calendar_events` and the **Put my follow-ups in my calendar** prompt: the client's
  own calendar tool (Claude with the Google Calendar connector, say) keeps a calendar in
  step with the tracker. Each event carries an id line, so the client finds it again to
  move or delete it instead of adding a second one, and never touches other events. The
  `.ics` export carries the same line.
- `add_column(options=[...])` adds a dropdown in Excel, as far down as the sheet's other
  dropdowns; a furthest-stage column (`field="reached"`) gets the stages by default. In
  Google Sheets the preview says to add the dropdown there.

## 0.4.3 — 2026-10-20

Changed
- Without a confirmation dialog, a preview now says, in words the model reads, to show
  it and end the turn, and to use the token only after the user's yes in a new message;
  it also carries the preview text to show. The evals caught a small model previewing
  and confirming in one turn (1 run of 6); with the new wording, 0 of 10.
- Evals print what the model did when a scenario fails, can repeat scenarios, and a
  failed eval fails the job.

## 0.4.2 — 2026-10-20

Added
- `add_column`: a new column after the last one, its header and cells styled like the
  neighbouring column (dates formatted like the submit dates), the filter extended.
  With `field` it is also mapped in the config — `replied_on`, `reached` or a name of
  your own — and the tools use it from the next call. Undo removes the column and the
  config line together. Works for Excel, CSV and Google Sheets (the sheet grows a
  column if it has none to spare).
- Evals for `check_tracker`, `reply_times`, `edit_application` and `add_column`. The
  evals workflow runs by itself on `main` when the code or the evals change, if the
  repository has an `ANTHROPIC_API_KEY` secret, and is skipped otherwise.

Fixed
- Two tests assumed the day they were written: the audit log keeps the real day of a
  write, and the tests now compare with it.

## 0.4.1 — 2026-10-19

Added
- `edit_application`: correct fields of one row by hand — a wrong or missing date, a
  typo in the company or role, the furthest stage, the link, a custom column — with
  the same preview, confirmation, backup and undo as the other writes. Dates are
  checked (a submit or reply date cannot be in the future, a reply cannot come before
  the submit) and written as dates; the reason and what changed go to the notes. The
  status stays with `set_status`, follow-ups with `log_followup` and
  `reschedule_followup`, which keep the related columns in step.

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
