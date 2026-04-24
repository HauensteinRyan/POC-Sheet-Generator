"""
Session auth for POC Sheet Generator.

Credentials are stored in users.json (gitignored, server-only).
Sessions use an HMAC-signed cookie — no DB required.
"""

import hashlib
import hmac
import json
import os
import secrets

USERS_PATH = os.path.join(os.path.dirname(__file__), "users.json")
SECRET_PATH = os.path.join(os.path.dirname(__file__), "secret.key")
COOKIE_NAME = "poc_session"


def _load_secret() -> str:
    """Load or generate the signing secret (persists across restarts)."""
    if os.path.exists(SECRET_PATH):
        with open(SECRET_PATH) as f:
            return f.read().strip()
    secret = secrets.token_hex(32)
    with open(SECRET_PATH, "w") as f:
        f.write(secret)
    return secret


SECRET = _load_secret()


def verify_credentials(username: str, password: str) -> bool:
    try:
        with open(USERS_PATH) as f:
            users = json.load(f)
    except Exception:
        return False
    stored = users.get(username)
    if stored is None:
        return False
    return secrets.compare_digest(str(stored), str(password))


def make_token(username: str) -> str:
    sig = hmac.new(SECRET.encode(), username.encode(), hashlib.sha256).hexdigest()
    return f"{sig}.{username}"


def verify_token(token: str | None) -> str | None:
    """Returns username if the token is valid, None otherwise."""
    if not token or "." not in token:
        return None
    sig, _, username = token.partition(".")
    expected = hmac.new(SECRET.encode(), username.encode(), hashlib.sha256).hexdigest()
    if secrets.compare_digest(sig, expected):
        return username
    return None
