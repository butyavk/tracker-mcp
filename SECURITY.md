# Security

## Reporting

Please report a vulnerability privately through GitHub (*Security* → *Report a
vulnerability* on this repository), not in a public issue.

## Threat model

tracker-mcp runs on your computer, started by your MCP client over stdio. It has no
network listener. What it can reach:

- **Your tracker** (`TRACKER_PATH`) and its config: read on every call, written only
  through the write tools.
- **Next to the tracker:** the config it writes on first start, and
  `backups-tracker-mcp/` with backups and `audit.jsonl` (or `TRACKER_BACKUP_DIR`).
- **Google Sheets**, only when the tracker is a Sheets URL: the Sheets API with the
  `spreadsheets` scope. With your own account that scope covers every spreadsheet
  you can open, although the server opens only the one in `TRACKER_PATH`. A service
  account limits it to the sheets you share with it. The sign-in is a refresh token
  in `~/.tracker-mcp/google-token.json` (readable only by you on macOS and Linux);
  treat it like a password.

No telemetry. The only outgoing connections are to Google, for a Sheets tracker.

### What it defends against

- **A model writing what you did not approve.** Every write needs your yes and the
  server checks it, not the model: a confirmation dialog (MCP elicitation) showing
  the exact cell changes, or, for clients without dialogs, a token from a dry run
  that fingerprints those changes and the tracker's version. What is written is
  exactly the approved plan, or nothing.
- **Prompt injection through your own data.** Notes, responses and logs are often
  pasted from job posts and e-mails. The server's instructions tell the model to
  treat cell text as data and never follow instructions in it, and the confirmation
  above stops an injected write even if the model is fooled. `evals/` includes a
  scenario with an injected note.
- **Formula injection.** Text the server writes stays text: `stringValue` in Google
  Sheets, a string cell in .xlsx, and in CSV a leading apostrophe for text that
  starts like a formula.
- **Lost data.** A backup before every write; writes refused while Excel has the
  file open, when the tracker changed since the read, or when openpyxl would drop
  parts of the workbook; every write verified before it replaces the original;
  undo only when nothing changed since.

### What it does not defend against

- A malicious MCP client or a compromised machine: the server acts with your user's
  file permissions.
- A model that shows you one preview and confirms another through the token path:
  the token proves a preview was made, not that you saw it. Clients with
  confirmation dialogs do not have this gap.
- Reading: read tools need no confirmation, so anything in the tracker can reach the
  model and, through it, other tools you have connected. Keep secrets out of the
  tracker.
