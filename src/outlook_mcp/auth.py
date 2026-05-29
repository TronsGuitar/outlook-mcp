"""OAuth2 (PKCE) against Microsoft's *consumer* endpoint + token storage/refresh.

Personal Microsoft accounts (outlook.com / hotmail / live) authenticate through the
`/consumers` authority. This is the key difference from work/school connectors that
reject personal accounts.

Two entry points:
  * `cli_login()`  - run once, interactively, to grant access and cache a refresh token.
  * `get_access_token()` - used by the server; silently refreshes using the cached token.
"""

from __future__ import annotations

import base64
import hashlib
import http.server
import json
import os
import secrets
import threading
import time
import urllib.parse
import webbrowser
from pathlib import Path

import httpx

# --- Configuration ---------------------------------------------------------

AUTHORITY = "https://login.microsoftonline.com/consumers/oauth2/v2.0"
AUTHORIZE_URL = f"{AUTHORITY}/authorize"
TOKEN_URL = f"{AUTHORITY}/token"

# Delegated Microsoft Graph scopes. offline_access is required to get a refresh token.
SCOPES = "openid profile offline_access https://graph.microsoft.com/Mail.Read"

TOKEN_PATH = Path(os.path.expanduser("~")) / ".outlook-mcp" / "token.json"


def get_client_id() -> str:
    client_id = os.environ.get("OUTLOOK_MCP_CLIENT_ID", "").strip()
    if not client_id:
        raise RuntimeError(
            "OUTLOOK_MCP_CLIENT_ID is not set. Set it to the Application (client) ID "
            "from your Azure app registration."
        )
    return client_id


def get_redirect_port() -> int:
    return int(os.environ.get("OUTLOOK_MCP_REDIRECT_PORT", "8765"))


def redirect_uri() -> str:
    return f"http://localhost:{get_redirect_port()}/callback"


# --- Token storage ---------------------------------------------------------


def _save_tokens(data: dict) -> None:
    TOKEN_PATH.parent.mkdir(parents=True, exist_ok=True)
    # Record absolute expiry so the server knows when to refresh.
    data = dict(data)
    data["expires_at"] = time.time() + int(data.get("expires_in", 3600)) - 60
    TOKEN_PATH.write_text(json.dumps(data, indent=2))
    try:
        os.chmod(TOKEN_PATH, 0o600)
    except OSError:
        pass


def _load_tokens() -> dict:
    if not TOKEN_PATH.exists():
        raise RuntimeError(
            f"No cached credentials at {TOKEN_PATH}. Run the one-time sign-in first:\n"
            f"  uv run outlook-mcp-auth"
        )
    return json.loads(TOKEN_PATH.read_text())


# --- PKCE helpers ----------------------------------------------------------


def _pkce_pair() -> tuple[str, str]:
    verifier = base64.urlsafe_b64encode(secrets.token_bytes(32)).rstrip(b"=").decode()
    challenge = (
        base64.urlsafe_b64encode(hashlib.sha256(verifier.encode()).digest())
        .rstrip(b"=")
        .decode()
    )
    return verifier, challenge


class _CallbackHandler(http.server.BaseHTTPRequestHandler):
    code: str | None = None
    error: str | None = None
    expected_state: str | None = None

    def do_GET(self):  # noqa: N802
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path != "/callback":
            self.send_response(404)
            self.end_headers()
            return
        params = urllib.parse.parse_qs(parsed.query)
        if params.get("state", [None])[0] != _CallbackHandler.expected_state:
            _CallbackHandler.error = "state_mismatch"
        elif "error" in params:
            _CallbackHandler.error = params.get("error_description", params["error"])[0]
        else:
            _CallbackHandler.code = params.get("code", [None])[0]
        self.send_response(200)
        self.send_header("Content-Type", "text/html")
        self.end_headers()
        msg = (
            "Sign-in complete. You can close this tab and return to Claude."
            if _CallbackHandler.code
            else f"Sign-in failed: {_CallbackHandler.error}"
        )
        self.wfile.write(f"<html><body><h2>{msg}</h2></body></html>".encode())

    def log_message(self, *args):  # silence the default stderr logging
        pass


def cli_login() -> None:
    """Interactive one-time login. Opens the browser, captures the redirect, caches tokens."""
    client_id = get_client_id()
    verifier, challenge = _pkce_pair()
    state = secrets.token_urlsafe(16)
    _CallbackHandler.code = None
    _CallbackHandler.error = None
    _CallbackHandler.expected_state = state

    params = {
        "client_id": client_id,
        "response_type": "code",
        "redirect_uri": redirect_uri(),
        "response_mode": "query",
        "scope": SCOPES,
        "state": state,
        "code_challenge": challenge,
        "code_challenge_method": "S256",
        "prompt": "select_account",
    }
    auth_url = f"{AUTHORIZE_URL}?{urllib.parse.urlencode(params)}"

    server = http.server.HTTPServer(("localhost", get_redirect_port()), _CallbackHandler)

    print("Opening your browser to sign in to Outlook...")
    print("If it does not open, paste this URL manually:\n")
    print(auth_url + "\n")
    try:
        webbrowser.open(auth_url)
    except Exception:
        pass

    # Serve requests until we get the callback (with a timeout safety net).
    server.timeout = 300
    deadline = time.time() + 300
    while _CallbackHandler.code is None and _CallbackHandler.error is None:
        if time.time() > deadline:
            raise RuntimeError("Timed out waiting for sign-in.")
        server.handle_request()

    if _CallbackHandler.error:
        raise RuntimeError(f"Authorization failed: {_CallbackHandler.error}")

    # Exchange the authorization code for tokens.
    resp = httpx.post(
        TOKEN_URL,
        data={
            "client_id": client_id,
            "grant_type": "authorization_code",
            "code": _CallbackHandler.code,
            "redirect_uri": redirect_uri(),
            "code_verifier": verifier,
            "scope": SCOPES,
        },
        timeout=30,
    )
    if resp.status_code != 200:
        raise RuntimeError(f"Token exchange failed ({resp.status_code}): {resp.text}")
    _save_tokens(resp.json())
    print(f"\nSuccess. Credentials cached to {TOKEN_PATH}")


def _refresh(tokens: dict) -> dict:
    refresh_token = tokens.get("refresh_token")
    if not refresh_token:
        raise RuntimeError("No refresh token cached. Run `uv run outlook-mcp-auth` again.")
    resp = httpx.post(
        TOKEN_URL,
        data={
            "client_id": get_client_id(),
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
            "scope": SCOPES,
        },
        timeout=30,
    )
    if resp.status_code != 200:
        raise RuntimeError(
            f"Token refresh failed ({resp.status_code}): {resp.text}. "
            "You may need to run `uv run outlook-mcp-auth` again."
        )
    new = resp.json()
    # Microsoft may not return a new refresh token; keep the old one if so.
    if "refresh_token" not in new:
        new["refresh_token"] = refresh_token
    _save_tokens(new)
    return new


def get_access_token() -> str:
    """Return a valid Graph access token, refreshing if needed."""
    tokens = _load_tokens()
    if time.time() >= tokens.get("expires_at", 0):
        tokens = _refresh(tokens)
    return tokens["access_token"]


if __name__ == "__main__":
    cli_login()
