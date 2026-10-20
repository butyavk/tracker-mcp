"""The few HTML pages a person sees on the remote server: the consent page during
sign-in (which client, which account, which tracker sheet), and plain notices."""

from __future__ import annotations

import html
from urllib.parse import urlparse

STYLE = """
:root { --bg:#f7f7f5; --fg:#1d1d1b; --muted:#5f5f5a; --card:#fff; --line:#e3e3df; --accent:#1f5fbf; }
@media (prefers-color-scheme: dark) {
  :root { --bg:#161615; --fg:#ececea; --muted:#a3a39e; --card:#1f1f1d; --line:#33332f; --accent:#7aa7ef; }
}
* { box-sizing: border-box; }
body { margin:0; background:var(--bg); color:var(--fg); font:16px/1.5 system-ui, -apple-system, sans-serif; }
main { max-width:560px; margin:48px auto; padding:0 16px; }
.card { background:var(--card); border:1px solid var(--line); border-radius:12px; padding:24px; }
h1 { font-size:22px; margin:0 0 12px; }
p { margin:8px 0; } .muted { color:var(--muted); font-size:14px; }
code { font-size:14px; }
a { color:var(--accent); }
.row { display:flex; flex-wrap:wrap; gap:8px; margin-top:20px; }
button, a.button { font:inherit; padding:10px 16px; border-radius:8px; border:1px solid var(--line);
  background:var(--card); color:var(--fg); cursor:pointer; text-decoration:none; }
button.primary { background:var(--accent); border-color:var(--accent); color:#fff; }
"""


def page(title: str, body: str, script: str = "") -> str:
    return (f"<!doctype html><html lang=en><head><meta charset=utf-8>"
            f"<meta name=viewport content='width=device-width, initial-scale=1'>"
            f"<title>{html.escape(title)}</title><style>{STYLE}</style></head>"
            f"<body><main><div class=card>{body}</div>"
            f"<p class=muted>tracker-mcp · <a href='https://github.com/butyavk/tracker-mcp'>source</a></p>"
            f"</main>{script}</body></html>")


def notice(title: str, text: str, back: str | None = None) -> str:
    safe = back and urlparse(back).scheme in ("https", "http")   # never a javascript: link
    link = f"<div class=row><a class=button href='{html.escape(back)}'>Back to the app</a></div>" if safe else ""
    return page(title, f"<h1>{html.escape(title)}</h1><p>{html.escape(text)}</p>{link}")


def consent(*, pending: str, client_name: str | None, redirect_uri: str, email: str,
            sheet_title: str | None, picker: bool) -> str:
    """Which client gets access, as whom, to which sheet. The client name is the client's
    own claim, so the address it returns to is shown just as plainly."""
    who = html.escape(client_name or "An MCP client")
    host = html.escape(urlparse(redirect_uri).netloc or redirect_uri)
    p = html.escape(pending)
    loop = (urlparse(redirect_uri).hostname or "") in ("localhost", "127.0.0.1", "::1", "[::1]")
    local = ("<p><b>That is a program on this computer</b> (Claude Code, the MCP Inspector). Any program on "
             "it could claim that name: continue only if you just started the connection there.</p>"
             if loop else "")
    if sheet_title:
        tracker = f"<p>Tracker: <b>{html.escape(sheet_title)}</b></p>"
        cont = "<button class=primary name=action value=continue>Continue</button>"
    else:
        tracker = "<p>No tracker yet: choose your sheet in Google Drive, or start a new one.</p>"
        cont = ""
    pick = ("<button type=button id=pick>Choose a sheet in Google Drive</button>" if picker else "")
    body = f"""
<h1>Connect your job tracker</h1>
<p><b>{who}</b> wants to read and update your job tracker as <b>{html.escape(email)}</b>.</p>
<p>It will send you back to <b>{host}</b>. Continue only if you started this yourself, from an app
at that address.</p>{local}
<p class=muted>Every change it makes still shows you a preview and needs your yes.</p>
{tracker}
<form method=post action=consent id=f>
<input type=hidden name=p value="{p}"><input type=hidden name=sheet_id id=sheet_id>
<div class=row>{cont}{pick}
<button name=action value=create>Start a new tracker sheet</button>
<button name=action value=cancel>Cancel</button></div></form>"""
    script = ""
    if picker:
        script = """<script src="https://apis.google.com/js/api.js"></script><script>
document.getElementById('pick').onclick = async () => {
  const p = document.querySelector('input[name=p]').value;
  const r = await fetch('consent/picker', {method: 'POST', body: new URLSearchParams({p})});
  const cfg = await r.json();
  if (!r.ok) { alert('This sign-in has expired: start again from your app.'); return; }
  gapi.load('picker', () => {
    const view = new google.picker.DocsView(google.picker.ViewId.SPREADSHEETS)
      .setMode(google.picker.DocsViewMode.LIST);
    new google.picker.PickerBuilder().addView(view).setOAuthToken(cfg.token).setDeveloperKey(cfg.key)
      .setAppId(cfg.app).setCallback(d => {
        if (d.action === google.picker.Action.PICKED) {
          document.getElementById('sheet_id').value = d.docs[0].id;
          const b = document.createElement('input'); b.type = 'hidden'; b.name = 'action'; b.value = 'pick';
          const f = document.getElementById('f'); f.appendChild(b); f.submit();
        }
      }).build().setVisible(true);
  });
};
</script>"""
    return page("Connect your job tracker", body, script)
