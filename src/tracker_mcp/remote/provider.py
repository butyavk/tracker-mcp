"""The remote server as its own OAuth authorization server for MCP clients, with Google
behind it for who the user is (the "third-party authorization" flow of the MCP spec).

    client ──/authorize──▶ /start (sets a cookie in this browser) ──▶ Google sign-in
                                                                          │
         /google/callback (same browser? allow-list; keep the Google token)
                                         ▼
           /consent: which client, as whom, which sheet ── Continue
                                         │
    client ◀── redirect_uri?code=… ──────┘      client ──/token (PKCE)──▶ server tokens

The sign-in is bound to the browser that started it: a cookie set on /start must come
back with Google's redirect and with the consent page, so someone who starts a sign-in
and sends the Google link to another person cannot finish it for them.

The tokens a client gets are the server's own, never Google's. Each sign-in starts a
token family: an access token for an hour and a refresh token rotated on every use, for
at most 90 days. Reusing a spent refresh token or authorization code revokes the whole
family (it means one was copied). Tokens are stored only as hashes.
"""

from __future__ import annotations

import ipaddress
import json
import secrets
import threading
import time
from collections import OrderedDict
from typing import Any
from urllib.parse import urlparse

import anyio.to_thread
from mcp.server.auth.provider import (
    AccessToken,
    AuthorizationCode,
    AuthorizationParams,
    AuthorizeError,
    OAuthAuthorizationServerProvider,
    RefreshToken,
    RegistrationError,
    TokenError,
    construct_redirect_uri,
)
from mcp.shared.auth import OAuthClientInformationFull, OAuthToken

from .storage import Storage, new_token, token_hash

SCOPE = "tracker"
ACCESS_TTL = 3600
REFRESH_TTL = 30 * 24 * 3600
FAMILY_TTL = 90 * 24 * 3600        # then the person signs in with Google again
CODE_TTL = 300
PENDING_TTL = 600
MAX_PENDING = 500
MAX_PENDING_PER_CLIENT = 20
MAX_CLIENTS = 500
UNUSED_CLIENT_TTL = 24 * 3600      # a registration never used for a token is dropped after a day
REPLAY_GRACE = 10                  # seconds: a refresh retried at once is not taken for a stolen copy
# exact https return addresses (Claude's connector callback); loopback ones are always allowed
DEFAULT_REDIRECTS = frozenset({"https://claude.ai/api/mcp/auth_callback",
                               "https://claude.com/api/mcp/auth_callback"})


def _is_loopback(host: str) -> bool:
    if host == "localhost":
        return True
    try:
        return ipaddress.ip_address(host.strip("[]")).is_loopback
    except ValueError:
        return False


def redirect_allowed(uri: str, allowed: frozenset[str]) -> bool:
    """An https address exactly on the list (or any https one, with "*"), or http on this
    machine on any port (a CLI or the MCP Inspector). Never javascript:, data: or other
    schemes, and no look-alike hosts, ports or paths."""
    u = urlparse(uri)
    host = (u.hostname or "").lower()
    if u.scheme == "http":
        return _is_loopback(host) and not u.username and not u.password
    if u.scheme != "https" or not host or u.username or u.password or u.fragment:
        return False
    return "*" in allowed or uri in allowed


class Provider(OAuthAuthorizationServerProvider[AuthorizationCode, RefreshToken, AccessToken]):
    def __init__(self, storage: Storage, base_url: str, resource: str,
                 allowed: frozenset[str] = frozenset(), redirects: frozenset[str] = DEFAULT_REDIRECTS):
        self.storage = storage
        self.base_url = base_url.rstrip("/")
        self.resource = resource
        self.allowed = allowed          # e-mails; someone taken off the list loses access at once
        self.redirects = redirects
        # sign-ins in progress live in memory (one instance): an unauthenticated /authorize
        # never writes to storage. Pages run in worker threads, hence the lock.
        self.pending: OrderedDict[str, dict[str, Any]] = OrderedDict()
        self._pending_lock = threading.Lock()
        # codes and refresh tokens are checked and spent under one lock (one instance), so
        # two requests cannot both redeem the same one
        self._grant_lock = threading.Lock()
        self._swept = 0.0

    # ---- helpers run in a worker thread: storage is a mounted bucket

    @staticmethod
    async def _io(fn, *args):
        return await anyio.to_thread.run_sync(fn, *args)

    def allows(self, subject: str | None) -> bool:
        user = self.storage.user(subject) if subject else None
        return bool(user and user.get("active", True) and user.get("email", "").lower() in self.allowed)

    def _family_ok(self, family: str | None) -> bool:
        data = self.storage.get("families", family) if family else None
        return bool(data and not data.get("revoked"))

    def revoke_family(self, family: str | None) -> None:
        if not family:
            return
        data = self.storage.get("families", family)
        if data:
            data["revoked"] = True
            self.storage.put("families", family, data)

    # ---- sign-ins in progress

    def new_pending(self, data: dict[str, Any], key: str | None = None) -> str:
        now = time.time()
        with self._pending_lock:
            for k in [k for k, v in self.pending.items() if v["expires_at"] < now]:
                del self.pending[k]
            mine = [k for k, v in self.pending.items() if v.get("client_id") == data.get("client_id")
                    and not v.get("cookie")]
            if len(mine) >= MAX_PENDING_PER_CLIENT:   # one client cannot crowd out the others
                del self.pending[mine[0]]
            while len(self.pending) >= MAX_PENDING:   # sign-ins nobody opened go first
                idle = next((k for k, v in self.pending.items() if not v.get("cookie")), None)
                del self.pending[idle if idle is not None else next(iter(self.pending))]
            key = key or new_token()
            self.pending[key] = {**data, "expires_at": now + PENDING_TTL}
            return key

    def get_pending(self, key: str) -> dict[str, Any] | None:
        with self._pending_lock:
            data = self.pending.get(key)
            if data and data["expires_at"] < time.time():
                self.pending.pop(key, None)
                return None
            return data

    def pop_pending(self, key: str) -> dict[str, Any] | None:
        with self._pending_lock:
            data = self.pending.pop(key, None)
        return data if data and data["expires_at"] >= time.time() else None

    # ---- clients (dynamic registration)

    async def get_client(self, client_id: str) -> OAuthClientInformationFull | None:
        def read():
            try:
                return self.storage.get("clients", client_id)
            except ValueError:
                return None
        data = await self._io(read)
        return OAuthClientInformationFull.model_validate(data) if data else None

    async def register_client(self, client_info: OAuthClientInformationFull) -> None:
        for uri in client_info.redirect_uris or []:
            if not redirect_allowed(str(uri), self.redirects):
                raise RegistrationError("invalid_redirect_uri", f"{uri} is not an address this server returns to")
        data = client_info.model_dump(mode="json", exclude_none=True)
        if len(json.dumps(data)) > 8192 or len(client_info.client_name or "") > 100:
            raise RegistrationError("invalid_client_metadata", "client metadata too large")
        # a registration that never gets a token expires; a used one is kept (see _issue)
        data["expires_at"] = int(time.time()) + UNUSED_CLIENT_TTL

        def write():
            folder = self.storage.root / "clients"
            files = sorted(folder.glob("*.json"), key=lambda p: p.stat().st_mtime) if folder.exists() else []
            if len(files) >= MAX_CLIENTS:   # make room: the oldest registration never used
                unused = next((p for p in files if (self.storage._read(p) or {}).get("expires_at")), None)
                if unused is None:
                    raise RegistrationError("invalid_client_metadata", "this server takes no more registrations")
                unused.unlink(missing_ok=True)
            self.storage.put("clients", client_info.client_id, data)
        await self._io(write)

    def _mark_used(self, client_id: str) -> None:
        data = self.storage.get("clients", client_id)
        if data and data.get("expires_at"):
            data.pop("expires_at")
            self.storage.put("clients", client_id, data)

    # ---- authorization: our /start, then Google, then /google/callback and /consent

    async def authorize(self, client: OAuthClientInformationFull, params: AuthorizationParams) -> str:
        if params.resource and params.resource.rstrip("/") != self.resource.rstrip("/"):
            raise AuthorizeError("invalid_target", f"this server issues tokens only for {self.resource}")
        if not redirect_allowed(str(params.redirect_uri), self.redirects):   # the list may have changed
            raise AuthorizeError("unauthorized_client", "this return address is no longer accepted")
        if time.time() - self._swept > 3600:   # expired codes, tokens and families, once an hour
            self._swept = time.time()
            await self._io(self.storage.sweep)
        key = self.new_pending({
            "client_id": client.client_id, "client_name": client.client_name,
            "redirect_uri": str(params.redirect_uri),
            "redirect_uri_provided_explicitly": params.redirect_uri_provided_explicitly,
            "state": params.state, "scopes": params.scopes or [SCOPE],
            "code_challenge": params.code_challenge, "resource": params.resource})
        return f"{self.base_url}/start?s={key}"

    def finish(self, pending: dict[str, Any], sub: str) -> str:
        """The user said Continue on the consent page: a code for the client."""
        code = new_token()
        self.storage.put("codes", token_hash(code), {
            "client_id": pending["client_id"], "scopes": pending["scopes"], "subject": sub,
            "code_challenge": pending["code_challenge"], "redirect_uri": pending["redirect_uri"],
            "redirect_uri_provided_explicitly": pending["redirect_uri_provided_explicitly"],
            "resource": pending.get("resource"), "expires_at": time.time() + CODE_TTL})
        return construct_redirect_uri(pending["redirect_uri"], code=code, state=pending.get("state"))

    @staticmethod
    def refuse(pending: dict[str, Any], error: str, description: str) -> str:
        return construct_redirect_uri(pending["redirect_uri"], error=error, error_description=description,
                                      state=pending.get("state"))

    async def load_authorization_code(self, client: OAuthClientInformationFull,
                                      authorization_code: str) -> AuthorizationCode | None:
        def read():
            data = self.storage.get("codes", token_hash(authorization_code))
            if not data or data["client_id"] != client.client_id:
                return None
            if data.get("used_family"):   # a second use: the code was copied; revoke what it gave
                self.revoke_family(data["used_family"])
                return None
            return data
        data = await self._io(read)
        if not data:
            return None
        return AuthorizationCode(code=authorization_code, scopes=data["scopes"], expires_at=data["expires_at"],
                                 client_id=data["client_id"], code_challenge=data["code_challenge"],
                                 redirect_uri=data["redirect_uri"],
                                 redirect_uri_provided_explicitly=data["redirect_uri_provided_explicitly"],
                                 resource=data.get("resource"), subject=data["subject"])

    async def exchange_authorization_code(self, client: OAuthClientInformationFull,
                                          authorization_code: AuthorizationCode) -> OAuthToken:
        def run():
            with self._grant_lock:
                h = token_hash(authorization_code.code)
                data = self.storage.get("codes", h)
                if not data or data.get("used_family"):
                    raise TokenError("invalid_grant", "authorization code already used or expired")
                family = new_token()
                now = int(time.time())
                self.storage.put("families", family, {"subject": authorization_code.subject,
                                                      "client_id": client.client_id, "created": now,
                                                      "expires_at": now + FAMILY_TTL})
                self.storage.put("codes", h, {**data, "used_family": family})   # kept until it expires
                self._mark_used(client.client_id)
                return self._issue(client.client_id, authorization_code.scopes, authorization_code.subject,
                                   authorization_code.resource, family)
        return await self._io(run)

    # ---- tokens

    def _issue(self, client_id: str, scopes: list[str], subject: str | None, resource: str | None,
               family: str) -> OAuthToken:
        if not self.allows(subject):
            raise TokenError("invalid_grant", "the user is not allowed on this server")
        fam = self.storage.get("families", family)
        if not fam or fam.get("revoked"):
            raise TokenError("invalid_grant", "this sign-in has ended; sign in again")
        access, refresh = new_token(), new_token()
        now = int(time.time())
        fam_end = int(fam["expires_at"])
        # a client that names no resource (RFC 8707) still gets a token bound to this server
        record = {"client_id": client_id, "scopes": scopes, "subject": subject,
                  "resource": resource or self.resource, "family": family}
        self.storage.put("access", token_hash(access), {**record, "expires_at": min(now + ACCESS_TTL, fam_end)})
        self.storage.put("refresh", token_hash(refresh), {**record, "expires_at": min(now + REFRESH_TTL, fam_end),
                                                          "access": token_hash(access)})
        return OAuthToken(access_token=access, token_type="Bearer", expires_in=max(1, min(ACCESS_TTL, fam_end - now)),
                          refresh_token=refresh, scope=" ".join(scopes))

    async def load_refresh_token(self, client: OAuthClientInformationFull, refresh_token: str) -> RefreshToken | None:
        def read():
            data = self.storage.get("refresh", token_hash(refresh_token))
            if not data or data["client_id"] != client.client_id:
                return None
            if data.get("spent"):   # a rotated token used again: a copy, unless a quick retry
                if time.time() - data.get("spent_at", 0) > REPLAY_GRACE:
                    self.revoke_family(data.get("family"))
                return None
            return data if self._family_ok(data.get("family")) else None
        data = await self._io(read)
        if not data:
            return None
        return RefreshToken(token=refresh_token, client_id=data["client_id"], scopes=data["scopes"],
                            expires_at=int(data["expires_at"]), resource=data.get("resource"),
                            subject=data["subject"])

    async def exchange_refresh_token(self, client: OAuthClientInformationFull, refresh_token: RefreshToken,
                                     scopes: list[str]) -> OAuthToken:
        def run():
            with self._grant_lock:
                h = token_hash(refresh_token.token)
                old = self.storage.get("refresh", h)
                if not old or old.get("spent"):
                    if old and time.time() - old.get("spent_at", 0) > REPLAY_GRACE:
                        self.revoke_family(old.get("family"))
                    raise TokenError("invalid_grant", "refresh token already used or expired")
                # kept, marked, to recognise a replay
                self.storage.put("refresh", h, {**old, "spent": True, "spent_at": time.time()})
                if old.get("access"):
                    self.storage.delete("access", old["access"])
                return self._issue(client.client_id, scopes or refresh_token.scopes, refresh_token.subject,
                                   refresh_token.resource, old["family"])
        return await self._io(run)

    async def load_access_token(self, token: str) -> AccessToken | None:
        def read():
            data = self.storage.get("access", token_hash(token))
            if not data or not self._family_ok(data.get("family")) or not self.allows(data.get("subject")):
                return None
            return data
        data = await self._io(read)
        if not data:
            return None
        return AccessToken(token=token, client_id=data["client_id"], scopes=data["scopes"],
                           expires_at=int(data["expires_at"]), resource=data.get("resource"),
                           subject=data["subject"])

    async def revoke_token(self, token: AccessToken | RefreshToken) -> None:
        def run():
            kind = "refresh" if isinstance(token, RefreshToken) else "access"
            data = self.storage.get(kind, token_hash(token.token))
            if data:
                self.revoke_family(data.get("family"))
        await self._io(run)


def new_cookie() -> str:
    return secrets.token_urlsafe(32)


__all__ = ["Provider", "SCOPE", "redirect_allowed", "new_cookie"]
