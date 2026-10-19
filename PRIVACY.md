# Privacy

tracker-mcp runs on your computer. It has no telemetry and no server of its own, and
it sends your data nowhere except where you point it:

- **Excel or CSV tracker:** the server reads and writes that file, its config next to
  it, and backups and an audit log in `backups-tracker-mcp/` next to it. Nothing
  leaves your computer through tracker-mcp. Your MCP client (Claude, Cursor, ...)
  sends what the tools return to its model, under that client's own privacy terms.
- **Google Sheets tracker:** the server talks to the Google Sheets API with your own
  Google sign-in, which is stored in `~/.tracker-mcp/google-token.json`. With
  `tracker-mcp auth google --pick` it can open only the sheets you pick in Google's
  file picker (scope `drive.file`). Backups of changed cells are kept in
  `~/.tracker-mcp/backups/`. Google's handling of that data is covered by
  https://policies.google.com/privacy.

To remove everything: delete `~/.tracker-mcp/` and the `backups-tracker-mcp/` folder,
and revoke tracker-mcp's access at https://myaccount.google.com/permissions if you
signed in to Google.

Questions: https://github.com/butyavk/tracker-mcp/issues
