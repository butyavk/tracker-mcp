"""What the remote server keeps between requests, as small JSON files under one folder
(on Cloud Run, a Cloud Storage bucket mounted as a folder).

    clients/<client id>.json        MCP clients that registered (RFC 7591)
    codes/<sha256>.json             authorization codes (5 minutes, used once)
    families/<id>.json              one sign-in's tokens (90 days at most; revoked as one)
    access/<sha256>.json            access tokens (1 hour)
    refresh/<sha256>.json           refresh tokens (30 days, rotated on use)
    users/<sub>/user.json           e-mail, Google refresh token (encrypted), the chosen sheet
    users/<sub>/tracker.tracker-mcp.toml, users/<sub>/backups/   the user's config and backups

Tokens are stored by their SHA-256, never as issued. The Google refresh token is
encrypted with a key that lives outside the folder (TRACKER_MCP_SECRET_KEY).
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import secrets
import time
from pathlib import Path
from typing import Any

from cryptography.fernet import Fernet, InvalidToken

KINDS = ("clients", "codes", "access", "refresh", "families")


def new_token() -> str:
    return secrets.token_urlsafe(32)   # 256 bits


def token_hash(token: str) -> str:
    return hashlib.sha256(token.encode()).hexdigest()


class Storage:
    def __init__(self, root: str | Path, secret_key: str):
        self.root = Path(root)
        self.root.mkdir(parents=True, exist_ok=True)
        self._fernet = Fernet(secret_key.encode() if isinstance(secret_key, str) else secret_key)

    # ---- files

    def _path(self, kind: str, name: str) -> Path:
        if kind not in KINDS or not re.fullmatch(r"[A-Za-z0-9_\-]{8,128}", name):
            raise ValueError("bad storage key")
        return self.root / kind / f"{name}.json"

    @staticmethod
    def _write(path: Path, data: dict[str, Any]) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp = path.with_name(f".{path.name}.{secrets.token_hex(4)}.tmp")
        tmp.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
        try:
            os.replace(tmp, path)
        except OSError:   # some mounted file systems cannot rename over a file
            path.write_text(tmp.read_text(encoding="utf-8"), encoding="utf-8")
            tmp.unlink(missing_ok=True)

    @staticmethod
    def _read(path: Path) -> dict[str, Any] | None:
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return None

    def put(self, kind: str, name: str, data: dict[str, Any]) -> None:
        self._write(self._path(kind, name), data)

    def get(self, kind: str, name: str) -> dict[str, Any] | None:
        """The record, or None if it is missing or expired (an expired one is removed)."""
        path = self._path(kind, name)
        data = self._read(path)
        if data and data.get("expires_at") and data["expires_at"] < time.time():
            path.unlink(missing_ok=True)
            return None
        return data

    def take(self, kind: str, name: str) -> dict[str, Any] | None:
        """Read and remove: for codes and pending sign-ins, which work once."""
        data = self.get(kind, name)
        self._path(kind, name).unlink(missing_ok=True)
        return data

    def delete(self, kind: str, name: str) -> None:
        self._path(kind, name).unlink(missing_ok=True)

    def sweep(self) -> int:
        """Remove expired records and stray temporary files; returns how many."""
        gone = 0
        now = time.time()
        for kind in KINDS:
            folder = self.root / kind
            for path in folder.glob("*.json"):
                data = self._read(path)
                if data and data.get("expires_at") and data["expires_at"] < now:
                    path.unlink(missing_ok=True)
                    gone += 1
            for path in folder.glob(".*.tmp"):   # left by an interrupted write
                try:
                    if path.stat().st_mtime < now - 3600:
                        path.unlink(missing_ok=True)
                        gone += 1
                except OSError:
                    pass
        return gone

    # ---- users

    def user_dir(self, sub: str) -> Path:
        if not re.fullmatch(r"[0-9A-Za-z_\-]{1,64}", sub):
            raise ValueError("bad user id")
        return self.root / "users" / sub

    def user(self, sub: str) -> dict[str, Any] | None:
        return self._read(self.user_dir(sub) / "user.json")

    def save_user(self, sub: str, data: dict[str, Any]) -> None:
        self._write(self.user_dir(sub) / "user.json", data)

    def encrypt(self, text: str) -> str:
        return self._fernet.encrypt(text.encode()).decode()

    def decrypt(self, text: str) -> str | None:
        try:
            return self._fernet.decrypt(text.encode()).decode()
        except (InvalidToken, ValueError):
            return None
