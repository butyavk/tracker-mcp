"""The remote server's side of Google: sign-in with a Web OAuth client, the user's
identity, a Google transport per user, and a new tracker sheet when the user has none.

Scopes: `openid email` (who signs in, checked against the allow-list) and `drive.file`
(only the spreadsheets the user picks or the server creates for them).
"""

from __future__ import annotations

import base64
import hashlib
import json
import secrets
import time
import urllib.parse
from dataclasses import dataclass
from typing import Any

from ..gsheets import API, TOKEN_URI, GoogleAuth, Transport, _post_form, http_transport
from ..scaffold import TEMPLATE_HEADERS, TEMPLATE_STATUSES
from ..stores import TrackerError

AUTH_URI = "https://accounts.google.com/o/oauth2/v2/auth"
SCOPES = "openid email https://www.googleapis.com/auth/drive.file"
ISSUERS = ("https://accounts.google.com", "accounts.google.com")


@dataclass(frozen=True)
class WebClient:
    client_id: str
    client_secret: str
    redirect_uri: str          # <base url>/google/callback


def pkce() -> tuple[str, str]:
    verifier = secrets.token_urlsafe(48)
    challenge = base64.urlsafe_b64encode(hashlib.sha256(verifier.encode()).digest()).rstrip(b"=").decode()
    return verifier, challenge


def authorize_url(client: WebClient, state: str, code_challenge: str, login_hint: str | None = None) -> str:
    params = {"client_id": client.client_id, "redirect_uri": client.redirect_uri, "response_type": "code",
              "scope": SCOPES, "access_type": "offline", "prompt": "consent", "state": state,
              "code_challenge": code_challenge, "code_challenge_method": "S256"}
    if login_hint:
        params["login_hint"] = login_hint
    return f"{AUTH_URI}?{urllib.parse.urlencode(params)}"


def exchange_code(client: WebClient, code: str, verifier: str) -> dict[str, Any]:
    return _post_form(TOKEN_URI, {"grant_type": "authorization_code", "code": code, "code_verifier": verifier,
                                  "client_id": client.client_id, "client_secret": client.client_secret,
                                  "redirect_uri": client.redirect_uri})


def identity(client: WebClient, tokens: dict[str, Any]) -> dict[str, Any]:
    """The signed-in user from the ID token. It came straight from Google's token endpoint
    over TLS, so its signature need not be checked (OpenID Connect Core 3.1.3.7); the
    audience, issuer, expiry and a verified e-mail still are."""
    raw = tokens.get("id_token") or ""
    try:
        payload = raw.split(".")[1]
        claims = json.loads(base64.urlsafe_b64decode(payload + "=" * (-len(payload) % 4)))
    except (IndexError, ValueError) as e:
        raise TrackerError("Google sent no readable ID token.") from e
    aud, cid = claims.get("aud"), client.client_id
    if not (aud == cid or (isinstance(aud, list) and cid in aud and claims.get("azp") == cid)) \
            or claims.get("iss") not in ISSUERS:
        raise TrackerError("The ID token is not for this server.")
    if int(claims.get("exp", 0)) < time.time():
        raise TrackerError("The ID token has expired.")
    if not claims.get("email") or claims.get("email_verified") is not True:
        raise TrackerError("Google did not confirm an e-mail address for this account.")
    return {"sub": str(claims["sub"]), "email": str(claims["email"]).lower()}


class UserGoogleAuth(GoogleAuth):
    """Access tokens for one user, from their refresh token (kept by the server)."""

    def __init__(self, client: WebClient, refresh_token: str, access: tuple[str, float] | None = None):
        super().__init__()
        self._info = {"refresh_token": refresh_token, "client_id": client.client_id,
                      "client_secret": client.client_secret, "token_uri": TOKEN_URI}
        self._token = access

    def _fetch(self) -> tuple[str, float]:
        return self._refresh(self._info)

    def access_token(self) -> str:
        return self.header().removeprefix("Bearer ")


def transport_for(auth: UserGoogleAuth) -> Transport:
    return http_transport(auth)


def create_tracker_sheet(call: Transport, title: str = "Job applications") -> str:
    """A new spreadsheet in the user's Drive with the `new` template's headers and a
    status dropdown; returns its id. drive.file covers files the app creates."""
    headers = [{"userEnteredValue": {"stringValue": h}, "userEnteredFormat": {"textFormat": {"bold": True}}}
               for h in TEMPLATE_HEADERS]
    body = {"properties": {"title": title},
            "sheets": [{"properties": {"title": "Applications", "gridProperties": {"frozenRowCount": 1}},
                        "data": [{"startRow": 0, "startColumn": 0, "rowData": [{"values": headers}]}]}]}
    made = call("POST", API, body)
    sid, sheet = made["spreadsheetId"], made["sheets"][0]["properties"]["sheetId"]
    status_col = TEMPLATE_HEADERS.index("Status")
    call("POST", f"{API}/{sid}:batchUpdate", {"requests": [{"setDataValidation": {
        "range": {"sheetId": sheet, "startRowIndex": 1, "endRowIndex": 1000,
                  "startColumnIndex": status_col, "endColumnIndex": status_col + 1},
        "rule": {"condition": {"type": "ONE_OF_LIST",
                               "values": [{"userEnteredValue": s} for s in TEMPLATE_STATUSES]},
                 "showCustomUi": True, "strict": False}}}]})
    return sid


def sheet_title(call: Transport, sid: str) -> str:
    try:
        return call("GET", f"{API}/{sid}?fields=properties.title", None)["properties"]["title"]
    except (TrackerError, KeyError):
        return sid
