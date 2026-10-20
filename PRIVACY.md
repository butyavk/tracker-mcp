# Privacy

tracker-mcp runs on your computer. It has no telemetry and the project runs no server,
and it sends your data nowhere except where you point it:

- **Excel or CSV tracker:** the server reads and writes that file, its config next to
  it, and backups and an audit log in `backups-tracker-mcp/` next to it. Nothing
  leaves your computer through tracker-mcp. Your MCP client (Claude, Cursor, ...)
  sends what the tools return to its model, under that client's own privacy terms.
- **Google Sheets tracker:** the server talks to the Google Sheets API with your own
  Google sign-in, which is stored in `~/.tracker-mcp/google-token.json`. With
  `tracker-mcp auth google` it can open only the sheets you pick in Google's file
  picker (scope `drive.file`). Backups of changed cells are kept in
  `~/.tracker-mcp/backups/`. tracker-mcp does not use data from Google for advertising
  or to train AI models and does not sell or share it; its use of Google user data
  follows the Google API Services User Data Policy, including the Limited Use
  requirements. Google's own handling of the data is covered by
  https://policies.google.com/privacy.

- **A remote tracker-mcp server** (`tracker-mcp remote`, used as a custom connector):
  it is run by a person, not by this project, in their own Google Cloud project. When
  you sign in to one, it keeps your e-mail, your Google sign-in (encrypted), the id of
  the sheet you picked, your config, backups of changed cells and an audit log of its
  writes. It reaches only the sheets you pick there. Its operator can see this data,
  so use a server run by someone you trust. To leave, remove the connector and
  tracker-mcp's access at https://myaccount.google.com/permissions, and ask the
  operator to delete your folder.

To remove everything: delete `~/.tracker-mcp/` and the `backups-tracker-mcp/` folder,
and revoke tracker-mcp's access at https://myaccount.google.com/permissions if you
signed in to Google.

Questions: https://github.com/butyavk/tracker-mcp/issues

The full policy, including what is read and written in a Google Sheet: https://butyavk.github.io/tracker-mcp/privacy.html
