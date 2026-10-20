"""The remote server as its own OAuth authorization server for MCP clients, with Google
behind it for who the user is (the "third-party authorization" flow of the MCP spec).

    client ──/authorize──▶ server ──▶ Google sign-in ──▶ /google/callback
                                         │ allow-list, keep the Google refresh token
                                         ▼
                           /consent: which client, which tracker sheet ── Continue
                                         │
    client ◀── redirect_uri?code=… ──────┘      client ──/token (PKCE)──▶ server tokens

The tokens a client gets are the server's own, never Google's: an access token for an
hour and a refresh token for 30 days, rotated on every use and stored only as hashes.
"""

from __future__ import annotations

import time
from typing import Any

from mcp.server.auth.provider import (
    AccessToken,
    AuthorizationCode,
    AuthorizationParams,
    AuthorizeError,
    OAuthAuthorizationServerProvider,
    RefreshToken,
    TokenError,
    construct_redirect_uri,
)
from mcp.shared.auth import OAuthClientInformationFull, OAuthToken

from . import google
from .storage import Storage, new_token, token_hash

SCOPE = "tracker"
ACCESS_TTL = 3600
REFRESH_TTL = 30 * 24 * 3600
CODE_TTL = 300
PENDING_TTL = 600


class Provider(OAuthAuthorizationServerProvider[AuthorizationCode, RefreshToken, AccessToken]):
    def __init__(self, storage: Storage, client: google.WebClient, base_url: str, resource: str,
                 allowed: frozenset[str] = frozenset()):
        self.storage = storage
        self.google = client
        self.base_url = base_url.rstrip("/")
        self.resource = resource
        self.allowed = allowed          # e-mails; someone taken off the list loses access at once
        self._swept = 0.0

    def allows(self, subject: str | None) -> bool:
        user = self.storage.user(subject) if subject else None
        return bool(user and user.get("active", True) and user.get("email", "").lower() in self.allowed)

    # ---- clients (dynamic registration)

    async def get_client(self, client_id: str) -> OAuthClientInformationFull | None:
        try:
            data = self.storage.get("clients", client_id)
        except ValueError:
            return None
        return OAuthClientInformationFull.model_validate(data) if data else None

    async def register_client(self, client_info: OAuthClientInformationFull) -> None:
        self.storage.put("clients", client_info.client_id, client_info.model_dump(mode="json", exclude_none=True))

    # ---- authorization: hand over to Google, come back to /google/callback

    async def authorize(self, client: OAuthClientInformationFull, params: AuthorizationParams) -> str:
        if params.resource and params.resource.rstrip("/") != self.resource.rstrip("/"):
            raise AuthorizeError("invalid_target", f"this server issues tokens only for {self.resource}")
        if time.time() - self._swept > 3600:   # expired codes, tokens and sign-ins, once an hour
            self._swept = time.time()
            self.storage.sweep()
        pending = new_token()
        verifier, challenge = google.pkce()
        self.storage.put("pending", pending, {
            "client_id": client.client_id, "client_name": client.client_name,
            "redirect_uri": str(params.redirect_uri),
            "redirect_uri_provided_explicitly": params.redirect_uri_provided_explicitly,
            "state": params.state, "scopes": params.scopes or [SCOPE],
            "code_challenge": params.code_challenge, "resource": params.resource,
            "google_verifier": verifier, "expires_at": time.time() + PENDING_TTL})
        return google.authorize_url(self.google, pending, challenge)

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
        data = self.storage.get("codes", token_hash(authorization_code))
        if not data or data["client_id"] != client.client_id:
            return None
        return AuthorizationCode(code=authorization_code, scopes=data["scopes"], expires_at=data["expires_at"],
                                 client_id=data["client_id"], code_challenge=data["code_challenge"],
                                 redirect_uri=data["redirect_uri"],
                                 redirect_uri_provided_explicitly=data["redirect_uri_provided_explicitly"],
                                 resource=data.get("resource"), subject=data["subject"])

    async def exchange_authorization_code(self, client: OAuthClientInformationFull,
                                          authorization_code: AuthorizationCode) -> OAuthToken:
        if not self.storage.take("codes", token_hash(authorization_code.code)):
            raise TokenError("invalid_grant", "authorization code already used or expired")
        return self._issue(client.client_id, authorization_code.scopes, authorization_code.subject,
                           authorization_code.resource)

    # ---- tokens

    def _issue(self, client_id: str, scopes: list[str], subject: str | None, resource: str | None) -> OAuthToken:
        if not self.allows(subject):
            raise TokenError("invalid_grant", "the user is not allowed on this server")
        access, refresh = new_token(), new_token()
        now = int(time.time())
        # a client that names no resource (RFC 8707) still gets a token bound to this server
        record = {"client_id": client_id, "scopes": scopes, "subject": subject, "resource": resource or self.resource}
        self.storage.put("access", token_hash(access), {**record, "expires_at": now + ACCESS_TTL,
                                                        "refresh": token_hash(refresh)})
        self.storage.put("refresh", token_hash(refresh), {**record, "expires_at": now + REFRESH_TTL,
                                                          "access": token_hash(access)})
        return OAuthToken(access_token=access, token_type="Bearer", expires_in=ACCESS_TTL,
                          refresh_token=refresh, scope=" ".join(scopes))

    async def load_refresh_token(self, client: OAuthClientInformationFull, refresh_token: str) -> RefreshToken | None:
        data = self.storage.get("refresh", token_hash(refresh_token))
        if not data or data["client_id"] != client.client_id:
            return None
        return RefreshToken(token=refresh_token, client_id=data["client_id"], scopes=data["scopes"],
                            expires_at=int(data["expires_at"]), resource=data.get("resource"),
                            subject=data["subject"])

    async def exchange_refresh_token(self, client: OAuthClientInformationFull, refresh_token: RefreshToken,
                                     scopes: list[str]) -> OAuthToken:
        old = self.storage.take("refresh", token_hash(refresh_token.token))   # rotation: the old one is spent
        if not old:
            raise TokenError("invalid_grant", "refresh token already used or expired")
        if old.get("access"):
            self.storage.delete("access", old["access"])
        return self._issue(client.client_id, scopes or refresh_token.scopes, refresh_token.subject,
                           refresh_token.resource)

    async def load_access_token(self, token: str) -> AccessToken | None:
        data = self.storage.get("access", token_hash(token))
        if not data or not self.allows(data.get("subject")):
            return None
        return AccessToken(token=token, client_id=data["client_id"], scopes=data["scopes"],
                           expires_at=int(data["expires_at"]), resource=data.get("resource"),
                           subject=data["subject"])

    async def revoke_token(self, token: AccessToken | RefreshToken) -> None:
        kind = "refresh" if isinstance(token, RefreshToken) else "access"
        data = self.storage.take(kind, token_hash(token.token))
        if data:
            other = data.get("access" if kind == "refresh" else "refresh")
            if other:
                self.storage.delete("access" if kind == "refresh" else "refresh", other)


__all__ = ["Provider", "SCOPE", "AuthorizeError"]
