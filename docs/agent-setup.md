# Setting up tracker-mcp: instructions for an AI agent

You are an agent (Claude Code, Codex or similar) asked to connect tracker-mcp to the
user's AI client. tracker-mcp is an MCP server for the spreadsheet where the user tracks
job applications (Excel, CSV or Google Sheets). Talk to the user in their language.
Humans: the same in Russian is at https://butyavk.github.io/tracker-mcp/ru.html.

## Rules

- Never edit the tracker file yourself, with a script or by hand. The server writes it,
  and only after the user says yes to a preview.
- Never overwrite a client's config: read it, back it up, add one entry, keep the rest.
- Google sign-in, passwords, OAuth secrets and API keys are the user's to type. Never ask
  for them in the chat.
- Ask before installing software or moving files.

## 1. Ask the user

1. Which client should use the tracker: Claude Desktop, Claude Code, Codex (CLI, IDE
   extension and app share one config), Cursor? Several is fine.
2. Where is the tracker: a path to an .xlsx or .csv file, a Google Sheets link, or none
   yet (then where to create one)?
3. Do they need it in claude.ai in the browser or on the phone? That needs their own
   server (section 5). Everything else runs on their computer. If they only want to try
   it there without running anything, tracker-mcp is not needed: claude.ai's Google
   Sheets connector and a project instruction will do (the Russian page has the text);
   say what they would miss (rules in code, enforced confirmation, backups and undo).

## 2. uv

Check `uvx --version`. If it is missing, install uv (it brings Python):

- macOS, Linux: `curl -LsSf https://astral.sh/uv/install.sh | sh`
- Windows (PowerShell): `powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"`

Then find the full path (`which uvx` / `where uvx`): clients started from the Dock or the
Start menu often do not see the shell's PATH.

## 3. The tracker

- **None yet:** `uvx tracker-mcp new <path>.xlsx` creates one with dropdowns, a filter, a
  summary sheet and its config.
- **Excel or CSV:** `uvx tracker-mcp init <path>` writes `<name>.tracker-mcp.toml` next to
  the tracker (it never touches the tracker) and lists what to check. Then
  `uvx tracker-mcp --check --tracker <path>` shows the mapping. Show it to the user and
  ask whether the columns and statuses are right. The config is plain TOML; edit it only
  if the user asks.
- **Google Sheets:** the user runs `uvx tracker-mcp auth google` in their own terminal
  (a browser opens for Google's sign-in and file picker; only the picked file is
  reachable). It writes the config and prints the TRACKER_PATH to use: the sheet's URL.

Running uvx here also caches the package, so the client's first start is fast.

## 4. Connect the client

Use the tracker's absolute path (or the sheet URL) as TRACKER_PATH.

- **Claude Code:**
  `claude mcp add tracker -s user -e TRACKER_PATH="<path>" -- uvx tracker-mcp`,
  then `claude mcp list` should show it connected. Or the plugin, which adds a weekly
  review skill: `/plugin install tracker-mcp --marketplace butyavk/tracker-mcp`.
- **Codex:** `codex mcp add tracker --env TRACKER_PATH="<path>" -- uvx tracker-mcp`, or
  add to `~/.codex/config.toml`:

  ```toml
  [mcp_servers.tracker]
  command = "uvx"            # full path if Codex does not find it
  args = ["tracker-mcp"]
  startup_timeout_sec = 30   # the first start may download packages

  [mcp_servers.tracker.env]
  TRACKER_PATH = "/absolute/path/job-tracker.xlsx"
  ```
- **Claude Desktop, Excel or CSV (preferred):** the extension. Get the `.mcpb` asset of
  the latest release (`https://api.github.com/repos/butyavk/tracker-mcp/releases/latest`,
  the `browser_download_url` ending in `.mcpb`), download it, and open it (`open <file>`
  on macOS). The user presses Install, picks the tracker in "Your tracker", Save, and
  switches the extension on. On macOS this is the only way that works with a tracker in
  Documents, Desktop or Downloads without more setup.
- **Claude Desktop, by hand** (Google Sheets, or a tracker outside those folders): add to
  `mcpServers` in `~/Library/Application Support/Claude/claude_desktop_config.json`
  (macOS) or `%APPDATA%\Claude\claude_desktop_config.json` (Windows):

  ```json
  "tracker": {
    "command": "/full/path/to/uvx",
    "args": ["tracker-mcp"],
    "env": { "TRACKER_PATH": "<path or sheet URL>" }
  }
  ```

  The user quits Claude Desktop completely (Cmd+Q on macOS) and opens it again.
- **Cursor:** the same entry in `~/.cursor/mcp.json`.

## 5. Claude on the web and the phone: the user's own server

Only for Google Sheets. The server runs in the user's own Google Cloud project; the
script `deploy/cloudrun.sh` in the repository does most of it. The user must do the
Google Cloud console steps in the browser: create the project, link billing, the
consent screen (External, then Publish app; scopes `openid`, `userinfo.email`,
`drive.file`), and an OAuth client of type Web application with the origin and redirect
URI that `./deploy/cloudrun.sh --project <id> --dry-run` prints. The deploy then runs in
Cloud Shell (or locally with gcloud signed in):

```sh
git clone -b remote-server https://github.com/butyavk/tracker-mcp && cd tracker-mcp
./deploy/cloudrun.sh --project <id> --client-id <web client id> --allow <their gmail>
```

The script asks for the client secret without showing it: the user pastes it into that
prompt, never into the chat. Afterwards: claude.ai → Customize → Connectors → Add custom
connector → `<server address>/mcp` → Connect, tick the Google Drive box, pick the sheet.
Every step in detail: https://butyavk.github.io/tracker-mcp/ru.html (section «Вариант 2»).

## 6. Check

In a new chat in the connected client, the user asks "describe my tracker". The answer
should list the columns and the statuses from the status dropdown. If the client shows
the tool as failed: Claude Desktop → Settings → Developer shows the server's log; the
usual causes are a wrong path, uvx not found (use the full path), or on macOS a tracker
in a protected folder (use the extension, or move the tracker).

Do not use one tracker through two connected clients in the same chat (for example the
local extension and a remote connector in Claude Desktop): the tools have the same
names and the model may call the wrong one. Do not keep an Excel file and its copy in
Google Sheets as two trackers: pick one.
