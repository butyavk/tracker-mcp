"""`tracker-mcp remote`: the same tools over HTTP, for many people, each signed in with
Google and working on their own Google Sheet. Settings come from the environment:

    TRACKER_MCP_BASE_URL         https://tracker-mcp-….run.app (no trailing slash)
    TRACKER_MCP_DATA_DIR         where to keep clients, tokens, users' configs and backups
    TRACKER_MCP_SECRET_KEY       a Fernet key: encrypts users' Google refresh tokens
    TRACKER_MCP_ALLOWED_EMAILS   who may sign in, comma-separated (required)
    GOOGLE_WEB_CLIENT_ID / GOOGLE_WEB_CLIENT_SECRET   a "Web application" OAuth client,
                                 redirect URI <base>/google/callback
    GOOGLE_PICKER_API_KEY, GOOGLE_APP_ID   optional: the Drive file picker on the consent
                                 page (the app id is the Cloud project number)
    TRACKER_MCP_REDIRECT_HOSTS   where MCP clients may be sent back after sign-in
                                 (default claude.ai,claude.com; loopback always)
    PORT                         default 8080
"""

from __future__ import annotations

import argparse
import hmac
import os
import re
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import anyio.to_thread
from mcp.server.auth.middleware.auth_context import get_access_token
from mcp.server.auth.provider import ProviderTokenVerifier
from mcp.server.auth.settings import AuthSettings, ClientRegistrationOptions, RevocationOptions
from mcp.server.mcpserver.exceptions import ToolError
from starlette.requests import Request
from starlette.responses import HTMLResponse, JSONResponse, PlainTextResponse, RedirectResponse, Response

from .. import server
from ..config import load_config_file
from ..gsheets import SheetsStore, use_transport
from ..stores import TrackerError
from ..tracker import Tracker
from . import google, pages
from .provider import DEFAULT_REDIRECT_HOSTS, SCOPE, Provider, new_cookie
from .storage import Storage, token_hash

CONFIG_NAME = "tracker.tracker-mcp.toml"
# The sign-in pages: no framing, no plugins, not cached; the referrer keeps only the
# origin (an API key restricted to this site needs it).
SECURITY_HEADERS = {"Cache-Control": "no-store", "X-Frame-Options": "DENY", "Referrer-Policy": "strict-origin",
                    "Content-Security-Policy": "frame-ancestors 'none'; object-src 'none'; base-uri 'none'",
                    "X-Content-Type-Options": "nosniff"}


@dataclass
class Settings:
    base_url: str
    data_dir: Path
    secret_key: str
    allowed: frozenset[str]
    client: google.WebClient
    picker_key: str | None = None
    app_id: str | None = None
    redirect_hosts: frozenset[str] = DEFAULT_REDIRECT_HOSTS

    @classmethod
    def from_env(cls, env: dict[str, str] | None = None) -> Settings:
        e = env if env is not None else dict(os.environ)
        missing = [k for k in ("TRACKER_MCP_BASE_URL", "TRACKER_MCP_SECRET_KEY", "TRACKER_MCP_ALLOWED_EMAILS",
                               "GOOGLE_WEB_CLIENT_ID", "GOOGLE_WEB_CLIENT_SECRET") if not e.get(k)]
        if missing:
            raise SystemExit(f"tracker-mcp remote: set {', '.join(missing)} (see `tracker-mcp remote --help`)")
        base = e["TRACKER_MCP_BASE_URL"].rstrip("/")
        allowed = frozenset(x.strip().lower() for x in e["TRACKER_MCP_ALLOWED_EMAILS"].split(",") if x.strip())
        return cls(base_url=base, data_dir=Path(e.get("TRACKER_MCP_DATA_DIR") or "/data"),
                   secret_key=e["TRACKER_MCP_SECRET_KEY"], allowed=allowed,
                   client=google.WebClient(e["GOOGLE_WEB_CLIENT_ID"], e["GOOGLE_WEB_CLIENT_SECRET"],
                                           f"{base}/google/callback"),
                   picker_key=e.get("GOOGLE_PICKER_API_KEY") or None, app_id=e.get("GOOGLE_APP_ID") or None,
                   redirect_hosts=frozenset(h.strip().lower() for h in e["TRACKER_MCP_REDIRECT_HOSTS"].split(",")
                                            if h.strip())
                   if e.get("TRACKER_MCP_REDIRECT_HOSTS") else DEFAULT_REDIRECT_HOSTS)


@dataclass
class Remote:
    settings: Settings
    storage: Storage
    provider: Provider
    _auth: dict[str, google.UserGoogleAuth] = field(default_factory=dict)
    _trackers: dict[str, tuple[tuple, Tracker]] = field(default_factory=dict)

    # ---- the signed-in user's Google access and tracker

    def google_auth(self, sub: str, user: dict[str, Any]) -> google.UserGoogleAuth:
        refresh = self.storage.decrypt(user.get("google", ""))
        if not refresh:
            raise ToolError("This server has no Google access for you any more: remove the connector and add "
                            "it again to sign in.")
        auth = self._auth.get(sub)
        if auth is None or auth._info["refresh_token"] != refresh:
            auth = self._auth[sub] = google.UserGoogleAuth(self.settings.client, refresh)
        return auth

    def tracker_for(self, sub: str) -> Tracker:
        user = self.storage.user(sub)
        if not user or not user.get("active", True):
            raise ToolError("You are not signed in to this server.")
        if not user.get("sheet"):
            raise ToolError("No tracker sheet chosen yet: remove the connector and add it again, then choose "
                            "or start a sheet on the sign-in page.")
        folder = self.storage.user_dir(sub)
        cfg_file = folder / CONFIG_NAME
        try:
            stat = cfg_file.stat()
        except OSError as e:
            raise ToolError("Your tracker's config is missing: reconnect to set it up again.") from e
        key = (user["sheet"], stat.st_mtime_ns, stat.st_size, user.get("google"))
        cached = self._trackers.get(sub)
        if cached and cached[0] == key:
            return cached[1]
        auth = self.google_auth(sub, user)
        store = SheetsStore(f"gsheet:{user['sheet']}", folder / "backups", transport=google.transport_for(auth))
        t = Tracker(f"gsheet:{user['sheet']}", config=load_config_file(cfg_file), store=store,
                    config_file=cfg_file)
        self._trackers[sub] = (key, t)
        return t

    def resolve(self) -> Tracker:
        token = get_access_token()
        if token is None or not token.subject:
            raise ToolError("Not signed in.")
        return self.tracker_for(token.subject)

    # ---- sign-in pages
    #
    # A sign-in is bound to the browser that started it: /start sets a cookie, and the
    # Google callback, the consent page and its form all need it. Whoever starts a
    # sign-in and sends the Google link to someone else cannot finish it for them.
    # Google calls block, so the pages do their work in a worker thread, as the SDK does
    # for the tools.

    @property
    def cookie_name(self) -> str:
        return "__Host-tmcp" if self.settings.base_url.startswith("https://") else "tmcp"

    def _cookie_ok(self, pending: dict[str, Any] | None, request: Request) -> bool:
        value = request.cookies.get(self.cookie_name, "")
        return bool(pending and value and pending.get("cookie")
                    and hmac.compare_digest(pending["cookie"], token_hash(value)))

    def _expired(self) -> Response:
        return HTMLResponse(pages.notice("Sign-in expired", "Start again from your app: this sign-in is older "
                                         "than ten minutes, was already used, or was started in another browser."),
                            status_code=400, headers=SECURITY_HEADERS)

    async def start(self, request: Request) -> Response:
        key = request.query_params.get("s", "")
        pending = self.provider.get_pending(key)
        if not pending or pending.get("cookie"):
            return self._expired()
        value = new_cookie()
        verifier, challenge = google.pkce()
        pending.update(cookie=token_hash(value), google_verifier=verifier)
        resp = RedirectResponse(google.authorize_url(self.settings.client, key, challenge), status_code=302)
        resp.set_cookie(self.cookie_name, value, max_age=600, path="/", httponly=True, samesite="lax",
                        secure=self.settings.base_url.startswith("https://"))
        return resp

    async def google_callback(self, request: Request) -> Response:
        q = dict(request.query_params)
        pending = self.provider.get_pending(q.get("state", ""))
        if not pending or pending.get("sub") or not pending.get("google_verifier"):
            return self._expired()
        if not self._cookie_ok(pending, request):
            return self._expired()
        self.provider.pop_pending(q["state"])
        if q.get("error") or not q.get("code"):
            return RedirectResponse(self.provider.refuse(pending, "access_denied", "Google sign-in was cancelled"),
                                    status_code=302)
        return await anyio.to_thread.run_sync(self._google_callback, q, pending)

    def _google_callback(self, q: dict[str, str], pending: dict[str, Any]) -> Response:
        try:
            tokens = google.exchange_code(self.settings.client, q["code"], pending["google_verifier"])
            who = google.identity(self.settings.client, tokens)
        except TrackerError as e:
            return HTMLResponse(pages.notice("Google sign-in failed", str(e)), status_code=400,
                                headers=SECURITY_HEADERS)
        if who["email"] not in self.settings.allowed:
            back = self.provider.refuse(pending, "access_denied", "this server is private")
            return HTMLResponse(pages.notice(
                "This server is private", f"{who['email']} is not on its list. Ask its owner to add you, or "
                "run tracker-mcp on your own computer.", back), status_code=403, headers=SECURITY_HEADERS)
        user = self.storage.user(who["sub"]) or {"created": int(time.time())}
        user.update(sub=who["sub"], email=who["email"], active=True, updated=int(time.time()))
        if tokens.get("refresh_token"):
            user["google"] = self.storage.encrypt(tokens["refresh_token"])
        if not user.get("google"):
            return HTMLResponse(pages.notice("Google gave no lasting access", "Remove tracker-mcp from your "
                                             "Google account's third-party access and connect again."),
                                status_code=400, headers=SECURITY_HEADERS)
        self.storage.save_user(who["sub"], user)
        self._trackers.pop(who["sub"], None)
        self._auth[who["sub"]] = google.UserGoogleAuth(
            self.settings.client, self.storage.decrypt(user["google"]) or "",
            (tokens["access_token"], time.time() + int(tokens.get("expires_in", 3600))))
        # a new id for the consent step: the Google state, seen in more places, is spent
        key = self.provider.new_pending({**pending, "sub": who["sub"], "email": who["email"]})
        return RedirectResponse(f"{self.settings.base_url}/consent?p={key}", status_code=302)

    def _consent_pending(self, p: str, request: Request) -> dict[str, Any] | None:
        pending = self.provider.get_pending(p)
        return pending if pending and pending.get("sub") and self._cookie_ok(pending, request) else None

    async def consent_page(self, request: Request) -> Response:
        p = request.query_params.get("p", "")
        pending = self._consent_pending(p, request)
        if not pending:
            return self._expired()
        return await anyio.to_thread.run_sync(self._consent_page, p, pending)

    def _consent_page(self, p: str, pending: dict[str, Any]) -> Response:
        user = self.storage.user(pending["sub"]) or {}
        title = None
        if user.get("sheet"):
            try:
                title = google.sheet_title(google.transport_for(self.google_auth(pending["sub"], user)), user["sheet"])
            except (TrackerError, ToolError):
                title = user["sheet"]
        return HTMLResponse(pages.consent(pending=p, client_name=pending.get("client_name"),
                                          redirect_uri=pending["redirect_uri"], email=pending["email"],
                                          sheet_title=title,
                                          picker=bool(self.settings.picker_key and self.settings.app_id)),
                            headers=SECURITY_HEADERS)

    async def picker_token(self, request: Request) -> Response:
        """The person's own Google token for the file picker, fetched by the consent page
        when they click Choose; never written into a page."""
        form = await request.form()
        pending = self._consent_pending(str(form.get("p", "")), request)
        if not pending or not (self.settings.picker_key and self.settings.app_id):
            return JSONResponse({"error": "expired"}, status_code=400)

        def token():
            user = self.storage.user(pending["sub"]) or {}
            return self.google_auth(pending["sub"], user).access_token()
        try:
            value = await anyio.to_thread.run_sync(token)
        except (TrackerError, ToolError) as e:
            return JSONResponse({"error": str(e)}, status_code=400)
        return JSONResponse({"token": value, "key": self.settings.picker_key, "app": self.settings.app_id},
                            headers={"Cache-Control": "no-store"})

    async def consent_submit(self, request: Request) -> Response:
        form = {k: str(v) for k, v in (await request.form()).items()}
        p = form.get("p", "")
        pending = self._consent_pending(p, request)
        if not pending:
            return self._expired()
        self.provider.pop_pending(p)   # one answer per consent page, even if the form is sent twice
        resp = await anyio.to_thread.run_sync(self._consent_submit, form, pending)
        if resp.status_code == 400:   # a failed pick or create: let them try again on the same page
            self.provider.pending[p] = pending
        else:
            resp.delete_cookie(self.cookie_name, path="/")
        return resp

    def _consent_submit(self, form: dict[str, str], pending: dict[str, Any]) -> Response:
        action = form.get("action", "")
        sub = pending["sub"]
        if action == "cancel":
            return RedirectResponse(self.provider.refuse(pending, "access_denied", "cancelled"), status_code=303)
        user = self.storage.user(sub) or {}
        try:
            call = google.transport_for(self.google_auth(sub, user))
            if action == "create":
                user["sheet"] = google.create_tracker_sheet(call)
            elif action == "pick":
                sid = form.get("sheet_id", "").strip()
                if not re.fullmatch(r"[A-Za-z0-9_-]{8,128}", sid):
                    raise TrackerError("No sheet was picked.")
                user["sheet"] = sid
            elif action != "continue" or not user.get("sheet"):
                raise TrackerError("Choose a sheet first.")
            if action in ("create", "pick"):
                self._write_config(sub, user["sheet"], call)
        except (TrackerError, ToolError) as e:
            return HTMLResponse(pages.notice("That did not work", str(e)), status_code=400,
                                headers=SECURITY_HEADERS)
        self.storage.save_user(sub, user)
        self._trackers.pop(sub, None)
        return RedirectResponse(self.provider.finish(pending, sub), status_code=303)

    def _write_config(self, sub: str, sid: str, call) -> None:
        """Read the sheet's headers and dropdowns and write the user's config, as
        `tracker-mcp init` does on a computer."""
        from ..scaffold import init
        folder = self.storage.user_dir(sub)
        with use_transport(call):
            init(f"gsheet:{sid}", out=folder / CONFIG_NAME, force=True)


def build(settings: Settings):
    """Wire the shared MCPServer for HTTP with sign-in, and return (remote, ASGI app)."""
    storage = Storage(settings.data_dir, settings.secret_key)
    resource = f"{settings.base_url}/mcp"
    provider = Provider(storage, settings.base_url, resource, settings.allowed, settings.redirect_hosts)
    remote = Remote(settings, storage, provider)
    mcp = server.mcp
    mcp.settings.auth = AuthSettings(
        issuer_url=settings.base_url, resource_server_url=resource, required_scopes=[SCOPE],
        client_registration_options=ClientRegistrationOptions(enabled=True, valid_scopes=[SCOPE],
                                                              default_scopes=[SCOPE]),
        revocation_options=RevocationOptions(enabled=True), validate_token_resource=True,
        service_documentation_url="https://github.com/butyavk/tracker-mcp#remote-server")
    mcp._auth_server_provider = provider
    mcp._token_verifier = ProviderTokenVerifier(provider)
    server._resolver = remote.resolve
    if not getattr(mcp, "_tracker_remote_routes", False):
        mcp.custom_route("/start", ["GET"])(_start)
        mcp.custom_route("/google/callback", ["GET"])(_google_callback)
        mcp.custom_route("/consent", ["GET"])(_consent_page)
        mcp.custom_route("/consent", ["POST"])(_consent_submit)
        mcp.custom_route("/consent/picker", ["POST"])(_picker_token)
        mcp.custom_route("/health", ["GET"])(_health)   # not /healthz: Cloud Run keeps that one
        mcp.custom_route("/", ["GET"])(_home)
        mcp._tracker_remote_routes = True
    remote_ref[0] = remote
    app = mcp.streamable_http_app(host="0.0.0.0")
    return remote, app


remote_ref: list[Any] = [None]   # the routes are registered once; tests build several servers


async def _start(request: Request) -> Response:
    return await remote_ref[0].start(request)


async def _picker_token(request: Request) -> Response:
    return await remote_ref[0].picker_token(request)


async def _google_callback(request: Request) -> Response:
    return await remote_ref[0].google_callback(request)


async def _consent_page(request: Request) -> Response:
    return await remote_ref[0].consent_page(request)


async def _consent_submit(request: Request) -> Response:
    return await remote_ref[0].consent_submit(request)


async def _health(request: Request) -> Response:
    return PlainTextResponse("ok")


async def _home(request: Request) -> Response:
    return HTMLResponse(pages.page("tracker-mcp", "<h1>tracker-mcp</h1><p>A job-application tracker for Claude "
                                   "and other MCP clients. Add <code>/mcp</code> on this address as a custom "
                                   "connector; you sign in with Google and pick your tracker sheet.</p>"))


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(prog="tracker-mcp remote", description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--port", type=int, default=int(os.environ.get("PORT", "8080")))
    p.add_argument("--host", default="0.0.0.0")
    a = p.parse_args(argv)
    import uvicorn

    _, app = build(Settings.from_env())
    print(f"tracker-mcp remote on {a.host}:{a.port}", file=sys.stderr)
    uvicorn.run(app, host=a.host, port=a.port, proxy_headers=True, forwarded_allow_ips="*", log_level="info")
    return 0
