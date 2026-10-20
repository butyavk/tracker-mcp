"""The remote server end to end, on a real HTTP port: an MCP client registers, the user
signs in with (a stand-in for) Google, picks a sheet on the consent page, the client
trades the code for tokens with PKCE and calls the tools; tokens rotate and expire, and
a stranger is turned away. Google itself is replaced by fakes."""

from __future__ import annotations

import base64
import hashlib
import json
import re
import secrets
import socket
import threading
import time
import urllib.parse
from datetime import datetime

import anyio
import httpx2
import pytest
from cryptography.fernet import Fernet
from mcp import Client
from mcp.client.streamable_http import streamable_http_client

from tracker_mcp import gsheets, server
from tracker_mcp.remote import app as remote_app
from tracker_mcp.remote import google
from tracker_mcp.remote.storage import Storage, token_hash

from .fake_sheets import FakeSheets

ALICE = {"sub": "1001", "email": "alice@example.com"}
BOB = {"sub": "1002", "email": "bob@example.com"}
EVE = {"sub": "1003", "email": "eve@example.com"}


def id_token(who: dict, aud: str = "web-client") -> str:
    claims = {"iss": "https://accounts.google.com", "aud": aud, "exp": int(time.time()) + 600,
              "email_verified": True, **who}
    body = base64.urlsafe_b64encode(json.dumps(claims).encode()).rstrip(b"=").decode()
    return f"e30.{body}.sig"


class Sheets:
    """One FakeSheets per spreadsheet id, plus the two calls FakeSheets does not know."""

    def __init__(self):
        self.books: dict[str, FakeSheets] = {}
        self.created = 0

    def add(self, sid: str, rows: list[list]) -> FakeSheets:
        self.books[sid] = FakeSheets({"Jobs": rows}, dropdowns={"Jobs": {3: [
            "Wishlist", "Applied", "Interview", "Offer", "Rejected", "Ghosted"]}})
        return self.books[sid]

    def __call__(self, method, url, body):
        if method == "POST" and url == gsheets.API:          # a new spreadsheet
            self.created += 1
            sid = f"new{self.created:04d}abcd"
            row = body["sheets"][0]["data"][0]["rowData"][0]["values"]
            headers = [v["userEnteredValue"]["stringValue"] for v in row]
            self.add(sid, [headers])
            return {"spreadsheetId": sid, "sheets": [{"properties": {"sheetId": 100}}]}
        sid = url.removeprefix(gsheets.API + "/").split("?")[0].split(":")[0]
        if url.endswith("?fields=properties.title"):
            return {"properties": {"title": f"Tracker {sid}"}}
        if method == "POST" and "setDataValidation" in json.dumps(body):
            return {}
        return self.books[sid](method, url, body)


@pytest.fixture
def live(tmp_path, monkeypatch):
    """A running remote server on a free port with fake Google; yields helpers."""
    sheets = Sheets()
    sheets.add("sheetALICE01", [["Company", "Role", "Status", "Date Applied", "Follow-up date", "Notes"],
                                ["Acme", "EM", "Applied", datetime(2026, 9, 1), datetime(2026, 10, 30), None]])
    sheets.add("sheetBOB0001", [["Company", "Role", "Status", "Date Applied", "Follow-up date", "Notes"],
                                ["Globex", "Director", "Interview", datetime(2026, 9, 2), None, None]])
    signing_in: dict[str, dict] = {}   # Google code -> who

    def exchange_code(client, code, verifier):
        assert verifier and client.redirect_uri.endswith("/google/callback")
        return {"access_token": f"g-{code}", "refresh_token": f"gr-{code}", "expires_in": 3600,
                "id_token": id_token(signing_in[code])}

    monkeypatch.setattr(google, "exchange_code", exchange_code)
    monkeypatch.setattr(google, "transport_for", lambda auth: sheets)
    monkeypatch.setattr(gsheets, "_post_form", lambda url, fields: {"access_token": "g-refreshed", "expires_in": 3600})
    monkeypatch.setenv("TRACKER_MCP_HOME", str(tmp_path / "home"))

    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        port = s.getsockname()[1]
    base = f"http://127.0.0.1:{port}"
    settings = remote_app.Settings.from_env({
        "TRACKER_MCP_BASE_URL": base, "TRACKER_MCP_DATA_DIR": str(tmp_path / "data"),
        "TRACKER_MCP_SECRET_KEY": Fernet.generate_key().decode(),
        "TRACKER_MCP_ALLOWED_EMAILS": "Alice@example.com, bob@example.com",
        "GOOGLE_WEB_CLIENT_ID": "web-client", "GOOGLE_WEB_CLIENT_SECRET": "web-secret",
        "GOOGLE_PICKER_API_KEY": "picker-key", "GOOGLE_APP_ID": "279023124405"})
    remote, app = remote_app.build(settings)
    import uvicorn
    srv = uvicorn.Server(uvicorn.Config(app, host="127.0.0.1", port=port, log_level="warning"))
    thread = threading.Thread(target=srv.run, daemon=True)
    thread.start()
    for _ in range(100):
        if srv.started:
            break
        time.sleep(0.05)

    http = httpx2.Client(base_url=base, follow_redirects=False, timeout=30)

    def register(redirect="http://127.0.0.1:9/cb", name="Test client"):
        r = http.post("/register", json={"redirect_uris": [redirect], "client_name": name,
                                         "token_endpoint_auth_method": "none",
                                         "grant_types": ["authorization_code", "refresh_token"],
                                         "response_types": ["code"]})
        assert r.status_code == 201, r.text
        return r.json()

    def sign_in(who, client, action="pick", sheet="sheetALICE01", redirect="http://127.0.0.1:9/cb"):
        verifier = secrets.token_urlsafe(48)
        challenge = base64.urlsafe_b64encode(hashlib.sha256(verifier.encode()).digest()).rstrip(b"=").decode()
        r = http.get("/authorize", params={
            "response_type": "code", "client_id": client["client_id"], "redirect_uri": redirect,
            "code_challenge": challenge, "code_challenge_method": "S256", "state": "st8",
            "scope": "tracker", "resource": f"{base}/mcp"})
        assert r.status_code == 302, r.text
        to_google = urllib.parse.urlparse(r.headers["location"])
        assert to_google.netloc == "accounts.google.com"
        q = dict(urllib.parse.parse_qsl(to_google.query))
        assert "drive.file" in q["scope"] and q["code_challenge_method"] == "S256"
        gcode = secrets.token_hex(8)
        signing_in[gcode] = who
        r = http.get("/google/callback", params={"code": gcode, "state": q["state"]})
        if r.status_code != 302:
            return r, None
        page = http.get(r.headers["location"])
        p = re.search(r'name=p value="([^"]+)"', page.text).group(1)
        r = http.post("/consent", data={"p": p, "action": action, "sheet_id": sheet})
        return page, (r, verifier)

    def tokens(client, r, verifier, redirect="http://127.0.0.1:9/cb"):
        assert r.status_code == 303, r.text
        back = urllib.parse.urlparse(r.headers["location"])
        q = dict(urllib.parse.parse_qsl(back.query))
        assert q["state"] == "st8"
        t = http.post("/token", data={"grant_type": "authorization_code", "code": q["code"],
                                      "redirect_uri": redirect, "client_id": client["client_id"],
                                      "code_verifier": verifier, "resource": f"{base}/mcp"})
        assert t.status_code == 200, t.text
        return t.json()

    async def call(token, tool, args=None):
        async with httpx2.AsyncClient(headers={"Authorization": f"Bearer {token}"}, timeout=30) as h:
            async with Client(streamable_http_client(f"{base}/mcp", http_client=h)) as c:
                return await c.call_tool(tool, args or {})

    yield type("Live", (), dict(base=base, http=http, remote=remote, sheets=sheets, register=staticmethod(register),
                                sign_in=staticmethod(sign_in), tokens=staticmethod(tokens),
                                call=staticmethod(call), data=tmp_path / "data"))
    srv.should_exit = True
    thread.join(5)
    server._resolver = None
    server.mcp.settings.auth = None
    server.mcp._auth_server_provider = None
    server.mcp._token_verifier = None


def test_discovery_and_no_token(live):
    r = live.http.get("/.well-known/oauth-protected-resource/mcp")
    assert r.json()["authorization_servers"] == [live.base] or r.json()["authorization_servers"] == [live.base + "/"]
    meta = live.http.get("/.well-known/oauth-authorization-server").json()
    assert meta["registration_endpoint"].endswith("/register") and "S256" in meta["code_challenge_methods_supported"]
    r = live.http.post("/mcp", json={"jsonrpc": "2.0", "id": 1, "method": "tools/list"},
                       headers={"Accept": "application/json, text/event-stream"})
    assert r.status_code == 401 and "resource_metadata" in r.headers["www-authenticate"]
    assert live.http.get("/healthz").text == "ok"


def test_sign_in_pick_a_sheet_and_use_the_tools(live):
    client = live.register()
    page, (r, verifier) = live.sign_in(ALICE, client)
    assert "Test client" in page.text and "127.0.0.1:9" in page.text and "alice@example.com" in page.text
    assert "picker-key" in page.text and "g-" in page.text      # the Drive picker gets the user's own token
    tok = live.tokens(client, r, verifier)
    assert tok["token_type"].lower() == "bearer" and tok["refresh_token"]

    async def go():
        d = (await live.call(tok["access_token"], "describe_tracker")).structured_content
        assert d["kind"] == "gsheet"
        prev = await live.call(tok["access_token"], "set_status", {"row": 2, "status": "Interview", "dry_run": True})
        token = prev.structured_content["confirm"]
        assert "end your turn" in prev.structured_content["next"]
        res = await live.call(tok["access_token"], "set_status", {"row": 2, "status": "Interview",
                                                                  "confirm": token})
        assert not res.is_error, res.content[0].text
        res = await live.call(tok["access_token"], "export_followups_calendar")
        assert res.is_error and "calendar_events" in res.content[0].text

    anyio.run(go)
    assert live.sheets.books["sheetALICE01"].value("Jobs", 2, 3) == "Interview"
    backups = list((live.data / "users" / "1001" / "backups").glob("*"))
    assert backups, "the write left a backup in the user's own folder"
    # stored by hash only; the Google refresh token encrypted
    raw = "".join(p.read_text() for p in live.data.rglob("*.json"))
    assert tok["access_token"] not in raw and tok["refresh_token"] not in raw and "gr-" not in raw


def test_refresh_rotates_and_revokes_the_old_pair(live):
    client = live.register()
    _, (r, verifier) = live.sign_in(ALICE, client)
    tok = live.tokens(client, r, verifier)
    new = live.http.post("/token", data={"grant_type": "refresh_token", "refresh_token": tok["refresh_token"],
                                         "client_id": client["client_id"]})
    assert new.status_code == 200, new.text
    again = live.http.post("/token", data={"grant_type": "refresh_token", "refresh_token": tok["refresh_token"],
                                           "client_id": client["client_id"]})
    assert again.status_code == 400 and again.json()["error"] == "invalid_grant"

    old = live.http.post("/mcp", json={"jsonrpc": "2.0", "id": 1, "method": "tools/list"},
                         headers={"Accept": "application/json, text/event-stream",
                                  "Authorization": f"Bearer {tok['access_token']}"})
    assert old.status_code == 401                                         # the old access token died with it

    async def go():
        d = (await live.call(new.json()["access_token"], "describe_tracker")).structured_content
        assert d["kind"] == "gsheet"

    anyio.run(go)


def test_a_stranger_is_turned_away_and_users_are_kept_apart(live):
    client = live.register()
    page, nothing = live.sign_in(EVE, client)
    assert page.status_code == 403 and "private" in page.text and nothing is None
    assert "error=access_denied" in page.text
    assert live.remote.storage.user("1003") is None

    _, (ra, va) = live.sign_in(ALICE, client)
    _, (rb, vb) = live.sign_in(BOB, client, sheet="sheetBOB0001")
    alice, bob = live.tokens(client, ra, va), live.tokens(client, rb, vb)

    async def go():
        a = (await live.call(alice["access_token"], "list_applications")).structured_content
        b = (await live.call(bob["access_token"], "list_applications")).structured_content
        assert [x["company"] for x in a["applications"]] == ["Acme"]
        assert [x["company"] for x in b["applications"]] == ["Globex"]

    anyio.run(go)
    # taken off the list: the next request is refused, and the refresh token is worthless
    live.remote.provider.allowed = frozenset({"alice@example.com"})
    r = live.http.post("/mcp", json={"jsonrpc": "2.0", "id": 1, "method": "tools/list"},
                       headers={"Accept": "application/json, text/event-stream",
                                "Authorization": f"Bearer {bob['access_token']}"})
    assert r.status_code == 401
    r = live.http.post("/token", data={"grant_type": "refresh_token", "refresh_token": bob["refresh_token"],
                                       "client_id": client["client_id"]})
    assert r.status_code == 400


def test_start_a_new_sheet_and_continue_next_time(live):
    client = live.register()
    _, (r, verifier) = live.sign_in(BOB, client, action="create")
    tok = live.tokens(client, r, verifier)
    assert live.remote.storage.user("1002")["sheet"] == "new0001abcd"

    async def go():
        d = (await live.call(tok["access_token"], "describe_tracker")).structured_content
        assert "Status" in json.dumps(d) or d["kind"] == "gsheet"

    anyio.run(go)
    page, (r, verifier) = live.sign_in(BOB, client, action="continue")
    assert "Tracker new0001abcd" in page.text                      # the sheet chosen last time
    live.tokens(client, r, verifier)


def test_codes_work_once_and_only_with_the_right_verifier(live):
    client = live.register()
    _, (r, verifier) = live.sign_in(ALICE, client)
    q = dict(urllib.parse.parse_qsl(urllib.parse.urlparse(r.headers["location"]).query))
    bad = live.http.post("/token", data={"grant_type": "authorization_code", "code": q["code"],
                                         "redirect_uri": "http://127.0.0.1:9/cb", "client_id": client["client_id"],
                                         "code_verifier": "x" * 50})
    assert bad.status_code == 400
    good = live.http.post("/token", data={"grant_type": "authorization_code", "code": q["code"],
                                          "redirect_uri": "http://127.0.0.1:9/cb",
                                          "client_id": client["client_id"], "code_verifier": verifier})
    twice = live.http.post("/token", data={"grant_type": "authorization_code", "code": q["code"],
                                           "redirect_uri": "http://127.0.0.1:9/cb",
                                           "client_id": client["client_id"], "code_verifier": verifier})
    assert good.status_code == 200, good.text           # a wrong verifier does not burn the code
    assert twice.status_code == 400


def test_storage_and_identity_checks(tmp_path):
    st = Storage(tmp_path, Fernet.generate_key().decode())
    st.put("codes", token_hash("c"), {"x": 1, "expires_at": time.time() - 1})
    assert st.get("codes", token_hash("c")) is None                 # expired
    with pytest.raises(ValueError):
        st.put("codes", "../../etc", {})
    enc = st.encrypt("secret")
    assert st.decrypt(enc) == "secret" and "secret" not in enc
    wc = google.WebClient("web-client", "s", "https://x/google/callback")
    assert google.identity(wc, {"id_token": id_token(ALICE)}) == ALICE
    with pytest.raises(Exception, match="not for this server"):
        google.identity(wc, {"id_token": id_token(ALICE, aud="someone-else")})
    with pytest.raises(Exception, match="e-mail"):
        google.identity(wc, {"id_token": id_token({"sub": "1", "email": "a@b.c", "email_verified": False})})
