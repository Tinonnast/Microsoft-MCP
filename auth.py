"""
Multi-user Microsoft OAuth (authorization code + PKCE) for personal accounts.

- Browser hits /login → Microsoft login → /auth/callback
- Server stores Graph tokens per Microsoft user id
- Issues a session token the MCP client must send as:
      Authorization: Bearer <session_token>
- After the first authenticated MCP request, the Mcp-Session-Id is bound to that
  user so follow-up posts in the same MCP session stay scoped correctly.
- A request with neither a valid Bearer nor a known Mcp-Session-Id is anonymous
  and cannot call Graph.
"""
from __future__ import annotations

import json
import os
import secrets
import threading
import time
from contextvars import ContextVar
from pathlib import Path
from typing import Any

import msal

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

CLIENT_ID = os.environ.get("MS_CLIENT_ID", "YOUR-CLIENT-ID")
CLIENT_SECRET = os.environ.get("MS_CLIENT_SECRET", "").strip() or None

AUTHORITY = os.environ.get(
    "MS_AUTHORITY", "https://login.microsoftonline.com/consumers"
)

REDIRECT_URI = os.environ.get(
    "MS_REDIRECT_URI", "http://localhost:8000/auth/callback"
)

PUBLIC_BASE_URL = os.environ.get("MS_PUBLIC_BASE_URL", "http://localhost:8000")

# Restricted scopes:
# - Files: read everywhere (Files.Read / Files.Read.All), write ONLY AppFolder
# - Mail: ReadBasic for listings; Mail.Read for full body (gated by sender PIN in code)
# - Calendar: ReadWrite so we can CREATE events (edit/delete blocked in code)
# - Tasks: ReadWrite for NEW tasks only (update/delete blocked in code)
# - Notes: read-only
# User must re-consent after changing these.
# offline_access / openid / profile are added by MSAL automatically — do not list them.
SCOPES = [
    "User.Read",
    "Files.Read",                 # read OneDrive (everywhere)
    "Files.Read.All",             # broader read if available (personal + shared)
    "Files.ReadWrite.AppFolder",  # write ONLY into the app's special folder
    "Notes.Read",
    "Tasks.ReadWrite",
    "Mail.ReadBasic",             # list/search without body
    "Mail.Read",                  # full body — only returned when PIN is correct
    "Calendars.ReadWrite",        # create events; edit/delete blocked in code
    "Contacts.Read",
]

# Always resolve data next to this package unless MS_DATA_DIR is absolute.
# Prevents "I deleted ./data" missing the real store because CWD differed.
_APP_DIR = Path(__file__).resolve().parent
_default_data = _APP_DIR / "data"
DATA_DIR = Path(os.environ.get("MS_DATA_DIR", str(_default_data))).expanduser()
if not DATA_DIR.is_absolute():
    DATA_DIR = (_APP_DIR / DATA_DIR).resolve()
else:
    DATA_DIR = DATA_DIR.resolve()
USERS_DIR = DATA_DIR / "users"
SESSIONS_FILE = DATA_DIR / "sessions.json"
PENDING_FILE = DATA_DIR / "pending_oauth.json"
MCP_BINDINGS_FILE = DATA_DIR / "mcp_session_bindings.json"

SESSION_TTL = int(os.environ.get("MS_SESSION_TTL", str(30 * 24 * 3600)))

current_access_token: ContextVar[str | None] = ContextVar(
    "current_access_token", default=None
)
current_user_id: ContextVar[str | None] = ContextVar("current_user_id", default=None)
current_user_name: ContextVar[str | None] = ContextVar(
    "current_user_name", default=None
)

# Thread-local backup (some SDK paths leave the event loop).
_tls = threading.local()
_lock = threading.Lock()


def _ensure_dirs() -> None:
    USERS_DIR.mkdir(parents=True, exist_ok=True)
    DATA_DIR.mkdir(parents=True, exist_ok=True)


def _load_json(path: Path, default: Any) -> Any:
    if not path.exists():
        return default
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return default


def _save_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, indent=2), encoding="utf-8")


# ---------------------------------------------------------------------------
# MSAL
# ---------------------------------------------------------------------------

def _build_msal_app(cache: msal.SerializableTokenCache | None = None) -> msal.ClientApplication:
    if CLIENT_ID == "YOUR-CLIENT-ID":
        raise RuntimeError(
            "Set MS_CLIENT_ID to your Entra app registration client ID."
        )
    kwargs: dict[str, Any] = {
        "client_id": CLIENT_ID,
        "authority": AUTHORITY,
        "token_cache": cache,
    }
    if CLIENT_SECRET:
        return msal.ConfidentialClientApplication(
            client_credential=CLIENT_SECRET, **kwargs
        )
    return msal.PublicClientApplication(**kwargs)


def _user_cache_path(user_id: str) -> Path:
    safe = "".join(c if c.isalnum() or c in "-_" else "_" for c in user_id)
    return USERS_DIR / f"{safe}.msal.json"


def _load_user_cache(user_id: str) -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    path = _user_cache_path(user_id)
    if path.exists():
        cache.deserialize(path.read_text(encoding="utf-8"))
    return cache


def _save_user_cache(user_id: str, cache: msal.SerializableTokenCache) -> None:
    if cache.has_state_changed:
        _user_cache_path(user_id).write_text(cache.serialize(), encoding="utf-8")


def _user_meta_path(user_id: str) -> Path:
    safe = "".join(c if c.isalnum() or c in "-_" else "_" for c in user_id)
    return USERS_DIR / f"{safe}.meta.json"


def save_user_meta(user_id: str, meta: dict) -> None:
    _ensure_dirs()
    _save_json(_user_meta_path(user_id), meta)


def load_user_meta(user_id: str) -> dict:
    return _load_json(_user_meta_path(user_id), {})


# ---------------------------------------------------------------------------
# Pending OAuth
# ---------------------------------------------------------------------------

def _load_pending() -> dict:
    return _load_json(PENDING_FILE, {})


def _save_pending(data: dict) -> None:
    _save_json(PENDING_FILE, data)


def start_login(extra: dict | None = None) -> str:
    """Begin auth-code + PKCE flow; return Microsoft authorize URL."""
    _ensure_dirs()
    cache = msal.SerializableTokenCache()
    app = _build_msal_app(cache=cache)
    flow = app.initiate_auth_code_flow(
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI,
    )
    if "auth_uri" not in flow:
        raise RuntimeError(f"Could not start login: {flow}")
    state = flow["state"]
    pending = _load_pending()
    entry: dict[str, Any] = {"flow": flow, "created_at": time.time()}
    if extra:
        entry.update(extra)
    pending[state] = entry
    # prune old pending (> 1 hour)
    now = time.time()
    pending = {k: v for k, v in pending.items() if now - v.get("created_at", 0) < 3600}
    _save_pending(pending)
    return flow["auth_uri"]


def finish_login(query_params: dict) -> dict:
    """Exchange code for tokens; return session_token + user info."""
    state = query_params.get("state")
    if not state:
        raise RuntimeError("Missing state in callback.")
    pending = _load_pending()
    entry = pending.pop(state, None)
    _save_pending(pending)
    if not entry:
        raise RuntimeError("Unknown or expired login state. Start again at /login.")

    flow = entry["flow"]
    cache = msal.SerializableTokenCache()
    app = _build_msal_app(cache=cache)

    result = app.acquire_token_by_auth_code_flow(flow, query_params)
    if "access_token" not in result:
        raise RuntimeError(
            result.get("error_description")
            or result.get("error")
            or str(result)
        )

    claims = result.get("id_token_claims") or {}
    user_id = claims.get("oid") or claims.get("sub")
    if not user_id:
        accounts = app.get_accounts()
        if accounts:
            user_id = accounts[0].get("home_account_id") or accounts[0].get(
                "local_account_id"
            )
    if not user_id:
        user_id = secrets.token_hex(16)

    display_name = claims.get("name") or claims.get("preferred_username") or "Microsoft user"
    email = claims.get("preferred_username") or claims.get("email") or ""

    _ensure_dirs()
    _save_user_cache(user_id, cache)
    save_user_meta(
        user_id,
        {
            "user_id": user_id,
            "display_name": display_name,
            "email": email,
            "last_login": time.time(),
        },
    )

    session_token = create_session(user_id)
    return {
        "session_token": session_token,
        "user_id": user_id,
        "display_name": display_name,
        "email": email,
    }


# ---------------------------------------------------------------------------
# Sessions (Bearer token → user_id)
# ---------------------------------------------------------------------------

def _load_sessions() -> dict:
    return _load_json(SESSIONS_FILE, {})


def _save_sessions(data: dict) -> None:
    _save_json(SESSIONS_FILE, data)


def create_session(user_id: str) -> str:
    _ensure_dirs()
    token = secrets.token_urlsafe(32)
    sessions = _load_sessions()
    sessions[token] = {
        "user_id": user_id,
        "created_at": time.time(),
        "expires_at": time.time() + SESSION_TTL,
    }
    now = time.time()
    sessions = {k: v for k, v in sessions.items() if v.get("expires_at", 0) > now}
    _save_sessions(sessions)
    return token


def resolve_session(session_token: str) -> dict | None:
    if not session_token:
        return None
    sessions = _load_sessions()
    entry = sessions.get(session_token)
    if not entry:
        return None
    if entry.get("expires_at", 0) < time.time():
        return None
    return entry


def revoke_session(session_token: str) -> None:
    sessions = _load_sessions()
    if session_token in sessions:
        del sessions[session_token]
        _save_sessions(sessions)


# ---------------------------------------------------------------------------
# MCP protocol session binding REMOVED (security).
# Previously, Mcp-Session-Id alone could restore another user's Graph token.
# Identity is now ONLY established from a valid Authorization: Bearer token
# on every request. No cross-request sticky identity without the secret token.
# ---------------------------------------------------------------------------

def bind_mcp_session(mcp_session_id: str, user_id: str) -> None:
    """No-op. Kept so older call sites do not crash. Never grants access."""
    return


def resolve_mcp_session(mcp_session_id: str) -> str | None:
    """Always None — session-id alone must never map to a user."""
    return None


# ---------------------------------------------------------------------------
# Graph access token
# ---------------------------------------------------------------------------

def get_access_token_for_user(user_id: str) -> str:
    """Acquire Graph token strictly for this user_id's own MSAL cache file."""
    if not user_id:
        raise RuntimeError("Refused: empty user_id")
    cache = _load_user_cache(user_id)
    app = _build_msal_app(cache=cache)
    accounts = app.get_accounts()
    result = None
    # Prefer account whose id matches this user_id; never borrow another cache.
    chosen = None
    for acc in accounts or []:
        hid = str(acc.get("home_account_id") or "")
        lid = str(acc.get("local_account_id") or "")
        if user_id in hid or user_id == lid or hid.startswith(user_id):
            chosen = acc
            break
    if chosen is None and accounts:
        # Single-account cache file for this user_id path is expected
        if len(accounts) == 1:
            chosen = accounts[0]
        else:
            raise RuntimeError(
                "Refused: ambiguous Microsoft accounts in user cache; re-login required."
            )
    if chosen is not None:
        result = app.acquire_token_silent(SCOPES, account=chosen)

    if not result or "access_token" not in result:
        raise RuntimeError(
            "Microsoft session expired or missing. Complete OAuth /login again."
        )

    _save_user_cache(user_id, cache)
    return result["access_token"]


def clear_request_user() -> None:
    """Wipe identity for this execution context. Safe to call often."""
    current_access_token.set(None)
    current_user_id.set(None)
    current_user_name.set(None)
    _tls.access_token = None
    _tls.user_id = None
    _tls.user_name = None


def _bind_user_id(user_id: str) -> dict:
    """Bind Graph credentials for exactly this user_id for the current request only."""
    if not user_id:
        raise PermissionError("Refusing to bind empty user_id.")
    token = get_access_token_for_user(user_id)
    if not token:
        raise PermissionError("No Graph token for user.")
    meta = load_user_meta(user_id)
    name = meta.get("display_name") or user_id
    # ContextVar only — request scoped. TLS mirror cleared/set but never read for auth.
    current_access_token.set(token)
    current_user_id.set(user_id)
    current_user_name.set(name)
    _tls.access_token = token
    _tls.user_id = user_id
    _tls.user_name = name
    return meta


def set_request_user(session_token: str, mcp_session_id: str | None = None) -> dict:
    """
    Validate Bearer session token and bind THAT user's Graph token only.
    mcp_session_id is ignored (never used for identity — prevents cross-user hijack).
    Raises PermissionError if the token is missing, unknown, or expired.
    """
    clear_request_user()
    if not session_token or not isinstance(session_token, str):
        raise PermissionError("Missing Bearer token.")
    session_token = session_token.strip()
    if not session_token:
        raise PermissionError("Missing Bearer token.")
    session = resolve_session(session_token)
    if not session:
        raise PermissionError(
            "Invalid or expired session token. Sign in again via OAuth or /login."
        )
    user_id = session.get("user_id")
    if not user_id:
        raise PermissionError("Corrupt session (no user_id).")
    # Only this user_id is ever bound for this request.
    return _bind_user_id(str(user_id))


def set_request_user_from_mcp_session(mcp_session_id: str) -> dict | None:
    """
    DISABLED for security. Mcp-Session-Id must never grant access.
    Always returns None after clearing identity.
    """
    clear_request_user()
    return None


def graph_headers() -> dict[str, str]:
    # Prefer ContextVar (request-scoped). Do not trust stale TLS across requests.
    token = current_access_token.get()
    uid = current_user_id.get()
    if not token or not uid:
        raise RuntimeError(
            "No authenticated user on this request. "
            "Send Authorization: Bearer <access_token> from OAuth /token or /login."
        )
    return {"Authorization": f"Bearer {token}"}


def require_user() -> None:
    uid = current_user_id.get()
    token = current_access_token.get()
    if not uid or not token:
        raise RuntimeError(
            "Not authenticated. Every request needs a valid Authorization: Bearer token "
            "issued to your account only. No Microsoft data is available without it."
        )


def whoami() -> str:
    uid = current_user_id.get()
    name = current_user_name.get()
    if not uid:
        return (
            "Not authenticated — no valid Authorization Bearer token on this request."
        )
    meta = load_user_meta(uid)
    email = meta.get("email") or ""
    return f"{name} <{email}> (id={uid})"


# ---------------------------------------------------------------------------
# MCP OAuth 2.1 Authorization Server (for ChatGPT / Grok / clients without
# custom headers). Issues the same session tokens used by the Bearer path.
# ---------------------------------------------------------------------------

OAUTH_CLIENTS_FILE = DATA_DIR / "oauth_clients.json"
OAUTH_CODES_FILE = DATA_DIR / "oauth_codes.json"
OAUTH_PENDING_FILE = DATA_DIR / "oauth_pending.json"
OAUTH_REFRESH_FILE = DATA_DIR / "oauth_refresh.json"

# Auth codes live 10 minutes; refresh tokens follow SESSION_TTL
AUTH_CODE_TTL = 600


def public_base() -> str:
    return PUBLIC_BASE_URL.rstrip("/")


def issuer() -> str:
    return public_base()


def resource_url() -> str:
    return public_base() + "/mcp"


def oauth_protected_resource_metadata() -> dict:
    """RFC 9728 — Protected Resource Metadata."""
    return {
        "resource": resource_url(),
        "authorization_servers": [issuer()],
        "bearer_methods_supported": ["header"],
        "scopes_supported": ["mcp"],
        "resource_documentation": public_base() + "/",
    }


def oauth_authorization_server_metadata() -> dict:
    """RFC 8414 — Authorization Server Metadata."""
    base = public_base()
    return {
        "issuer": issuer(),
        "authorization_endpoint": base + "/authorize",
        "token_endpoint": base + "/token",
        "registration_endpoint": base + "/register",
        "scopes_supported": ["mcp", "openid"],
        "response_types_supported": ["code"],
        "grant_types_supported": ["authorization_code", "refresh_token"],
        "token_endpoint_auth_methods_supported": [
            "none",
            "client_secret_post",
            "client_secret_basic",
        ],
        "code_challenge_methods_supported": ["S256", "plain"],
        "client_id_metadata_document_supported": True,
        "service_documentation": base + "/",
    }


def _load_oauth_clients() -> dict:
    return _load_json(OAUTH_CLIENTS_FILE, {})


def _save_oauth_clients(data: dict) -> None:
    _save_json(OAUTH_CLIENTS_FILE, data)


def _load_oauth_codes() -> dict:
    return _load_json(OAUTH_CODES_FILE, {})


def _save_oauth_codes(data: dict) -> None:
    _save_json(OAUTH_CODES_FILE, data)


def _load_oauth_pending() -> dict:
    return _load_json(OAUTH_PENDING_FILE, {})


def _save_oauth_pending(data: dict) -> None:
    _save_json(OAUTH_PENDING_FILE, data)


def _load_oauth_refresh() -> dict:
    return _load_json(OAUTH_REFRESH_FILE, {})


def _save_oauth_refresh(data: dict) -> None:
    _save_json(OAUTH_REFRESH_FILE, data)


def register_oauth_client(meta: dict) -> dict:
    """
    Dynamic Client Registration (RFC 7591).
    Also accepts clients whose client_id is a URL (CIMD) without registration.
    """
    _ensure_dirs()
    redirect_uris = meta.get("redirect_uris") or []
    if not isinstance(redirect_uris, list) or not redirect_uris:
        raise ValueError("redirect_uris required")

    client_id = secrets.token_urlsafe(24)
    client_secret = None
    # Public clients (token_endpoint_auth_method=none) get no secret
    auth_method = meta.get("token_endpoint_auth_method") or "none"
    if auth_method not in ("none", "client_secret_post", "client_secret_basic"):
        auth_method = "none"
    if auth_method != "none":
        client_secret = secrets.token_urlsafe(32)

    entry = {
        "client_id": client_id,
        "client_secret": client_secret,
        "redirect_uris": list(redirect_uris),
        "grant_types": meta.get("grant_types")
        or ["authorization_code", "refresh_token"],
        "response_types": meta.get("response_types") or ["code"],
        "token_endpoint_auth_method": auth_method,
        "client_name": meta.get("client_name") or "",
        "created_at": time.time(),
    }
    clients = _load_oauth_clients()
    clients[client_id] = entry
    _save_oauth_clients(clients)

    out = {
        "client_id": client_id,
        "redirect_uris": entry["redirect_uris"],
        "grant_types": entry["grant_types"],
        "response_types": entry["response_types"],
        "token_endpoint_auth_method": auth_method,
        "client_id_issued_at": int(entry["created_at"]),
    }
    if client_secret:
        out["client_secret"] = client_secret
    return out


def get_oauth_client(client_id: str) -> dict | None:
    if not client_id:
        return None
    clients = _load_oauth_clients()
    if client_id in clients:
        return clients[client_id]
    # Client ID Metadata Document (CIMD): client_id is an https URL.
    # We accept it without pre-registration and trust redirect_uri matching later.
    if client_id.startswith("https://"):
        return {
            "client_id": client_id,
            "client_secret": None,
            "redirect_uris": None,  # open — validated against request only for CIMD
            "grant_types": ["authorization_code", "refresh_token"],
            "response_types": ["code"],
            "token_endpoint_auth_method": "none",
            "cimd": True,
        }
    return None


def start_oauth_authorize(
    *,
    client_id: str,
    redirect_uri: str,
    state: str,
    code_challenge: str | None,
    code_challenge_method: str | None,
    scope: str | None = None,
) -> str:
    """
    Validate MCP client OAuth request, store it, start Microsoft login.
    Returns Microsoft authorize URL.
    """
    client = get_oauth_client(client_id)
    if not client:
        raise ValueError(
            "Unknown client_id. Register via POST /register, or use a CIMD https URL as client_id."
        )
    allowed = client.get("redirect_uris")
    if allowed is not None and redirect_uri not in allowed:
        raise ValueError("redirect_uri is not registered for this client")
    if (
        not redirect_uri.startswith("https://")
        and not redirect_uri.startswith("http://localhost")
        and not redirect_uri.startswith("http://127.0.0.1")
    ):
        raise ValueError("redirect_uri must be https (or localhost for dev)")

    method = (code_challenge_method or "plain").upper()
    if code_challenge and method not in ("S256", "PLAIN"):
        raise ValueError("code_challenge_method must be S256 or plain")

    bridge = secrets.token_urlsafe(24)
    oauth_pending = _load_oauth_pending()
    oauth_pending[bridge] = {
        "client_id": client_id,
        "redirect_uri": redirect_uri,
        "state": state or "",
        "code_challenge": code_challenge or "",
        "code_challenge_method": method,
        "scope": scope or "mcp",
        "created_at": time.time(),
    }
    now = time.time()
    oauth_pending = {
        k: v for k, v in oauth_pending.items() if now - v.get("created_at", 0) < 3600
    }
    _save_oauth_pending(oauth_pending)

    # Microsoft login stores oauth_bridge on the pending MSAL state
    ms_url = start_login(extra={"oauth_bridge": bridge})
    return ms_url


def complete_oauth_after_microsoft(query_params: dict) -> dict:
    """
    Called from /auth/callback. Finishes Microsoft login, and if this was an
    MCP OAuth authorize flow, returns redirect info for the AI client.
    Otherwise returns the normal session success payload.
    """
    state = query_params.get("state")
    pending_ms = _load_pending()
    entry = pending_ms.get(state or "") if state else None
    bridge = (entry or {}).get("oauth_bridge") if entry else None

    info = finish_login(query_params)  # creates session + MS tokens

    if not bridge:
        return {"mode": "manual", **info}

    oauth_pending = _load_oauth_pending()
    req = oauth_pending.pop(bridge, None)
    _save_oauth_pending(oauth_pending)
    if not req:
        return {"mode": "manual", **info}

    # Issue one-time authorization code bound to this user session token
    code = secrets.token_urlsafe(32)
    codes = _load_oauth_codes()
    codes[code] = {
        "client_id": req["client_id"],
        "redirect_uri": req["redirect_uri"],
        "code_challenge": req.get("code_challenge") or "",
        "code_challenge_method": req.get("code_challenge_method") or "S256",
        "session_token": info["session_token"],
        "user_id": info["user_id"],
        "scope": req.get("scope") or "mcp",
        "expires_at": time.time() + AUTH_CODE_TTL,
        "created_at": time.time(),
    }
    now = time.time()
    codes = {k: v for k, v in codes.items() if v.get("expires_at", 0) > now}
    _save_oauth_codes(codes)

    from urllib.parse import urlencode

    q = {"code": code}
    if req.get("state"):
        q["state"] = req["state"]
    redirect_to = req["redirect_uri"]
    sep = "&" if ("?" in redirect_to) else "?"
    return {
        "mode": "oauth_redirect",
        "redirect_to": redirect_to + sep + urlencode(q),
        "display_name": info.get("display_name"),
        "email": info.get("email"),
    }


def _verify_pkce(code_verifier: str, challenge: str, method: str) -> bool:
    if not challenge:
        return True  # client did not use PKCE (allowed for confidential)
    if not code_verifier:
        return False
    method = (method or "S256").upper()
    if method == "PLAIN":
        return secrets.compare_digest(code_verifier, challenge)
    # S256
    import hashlib
    import base64

    digest = hashlib.sha256(code_verifier.encode("ascii")).digest()
    computed = base64.urlsafe_b64encode(digest).rstrip(b"=").decode("ascii")
    return secrets.compare_digest(computed, challenge)


def exchange_oauth_token(
    *,
    grant_type: str,
    code: str | None = None,
    redirect_uri: str | None = None,
    client_id: str | None = None,
    client_secret: str | None = None,
    code_verifier: str | None = None,
    refresh_token: str | None = None,
) -> dict:
    """Token endpoint handler body."""
    if grant_type == "authorization_code":
        if not code:
            raise ValueError("code required")
        codes = _load_oauth_codes()
        entry = codes.pop(code, None)
        _save_oauth_codes(codes)
        if not entry:
            raise PermissionError("invalid_grant: code unknown or already used")
        if entry.get("expires_at", 0) < time.time():
            raise PermissionError("invalid_grant: code expired")
        if redirect_uri and redirect_uri != entry.get("redirect_uri"):
            raise PermissionError("invalid_grant: redirect_uri mismatch")
        if client_id and client_id != entry.get("client_id"):
            raise PermissionError("invalid_client: client_id mismatch")

        client = get_oauth_client(entry["client_id"])
        if client and client.get("client_secret"):
            if not client_secret or not secrets.compare_digest(
                client_secret, client["client_secret"]
            ):
                raise PermissionError("invalid_client: bad client_secret")

        if not _verify_pkce(
            code_verifier or "",
            entry.get("code_challenge") or "",
            entry.get("code_challenge_method") or "S256",
        ):
            raise PermissionError("invalid_grant: PKCE verification failed")

        access = entry["session_token"]
        # Issue refresh token
        rt = secrets.token_urlsafe(32)
        refreshes = _load_oauth_refresh()
        refreshes[rt] = {
            "session_token": access,
            "user_id": entry["user_id"],
            "client_id": entry["client_id"],
            "expires_at": time.time() + SESSION_TTL,
            "created_at": time.time(),
        }
        _save_oauth_refresh(refreshes)

        return {
            "access_token": access,
            "token_type": "Bearer",
            "expires_in": SESSION_TTL,
            "refresh_token": rt,
            "scope": entry.get("scope") or "mcp",
        }

    if grant_type == "refresh_token":
        if not refresh_token:
            raise ValueError("refresh_token required")
        refreshes = _load_oauth_refresh()
        entry = refreshes.get(refresh_token)
        if not entry or entry.get("expires_at", 0) < time.time():
            raise PermissionError("invalid_grant: refresh_token invalid or expired")
        if client_id and client_id != entry.get("client_id"):
            raise PermissionError("invalid_client")

        # Rotate access token (new session for same user)
        new_access = create_session(entry["user_id"])
        entry["session_token"] = new_access
        entry["expires_at"] = time.time() + SESSION_TTL
        refreshes[refresh_token] = entry
        _save_oauth_refresh(refreshes)

        return {
            "access_token": new_access,
            "token_type": "Bearer",
            "expires_in": SESSION_TTL,
            "refresh_token": refresh_token,
            "scope": "mcp",
        }

    raise ValueError(f"unsupported_grant_type: {grant_type}")
