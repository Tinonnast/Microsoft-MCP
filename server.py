#!/usr/bin/env python3
"""
Microsoft Graph MCP server — multi-user, mostly read-only.

Allowed writes (only):
  - Create NEW files under OneDrive /AI (common doc/text types; no overwrite)
  - Create NEW To Do tasks (no complete/update/delete, no new lists)

Everything else is read-only: no deletes, no updates of existing files,
no OneNote create/edit, no sending mail, no calendar writes.

Each user:
  1. Opens https://<your-server>/login in a browser
  2. Signs in with Microsoft (personal account)
  3. Copies the session token shown on the success page
  4. Configures their MCP client with:
       URL:  https://<your-server>/mcp
       Header: Authorization: Bearer <session_token>
"""
from __future__ import annotations

import os
from typing import Any

from mcp.server.mcpserver import MCPServer
from mcp.server.transport_security import TransportSecuritySettings
from starlette.applications import Starlette
from starlette.requests import Request
from starlette.responses import HTMLResponse, JSONResponse, Response
from starlette.routing import Mount, Route
from starlette.types import ASGIApp, Receive, Scope, Send

import auth
import graph

# ---------------------------------------------------------------------------
# MCP tools
# ---------------------------------------------------------------------------

mcp = MCPServer(
    "microsoft-graph-personal",
    instructions=(
        "Access is limited strictly to the currently authenticated user's own Microsoft "
        "account. Unauthenticated or cross-account access is impossible. "
        "Read: OneDrive, OneNote, To Do, Calendar, Contacts, email metadata. "
        "Full email body only via read_email with a valid pin. "
        "Write only: new files in AppFolder, new To Do tasks, new calendar events. "
        "Never delete, overwrite, send mail, edit OneNote, or edit/delete tasks or events."
    ),
)


@mcp.tool()
async def whoami() -> str:
    """Show which Microsoft account this session is using."""
    return auth.whoami()


# ---- OneDrive ----

@mcp.tool()
async def search_onedrive(query: str, top: int = 40) -> str:
    """Search OneDrive by filename or content (paginated). Use for names, stems, or extensions like pptx."""
    auth.require_user()
    hits = graph.search_files(query, top=top)
    if not hits:
        return "No files found. Try a shorter stem (without path), or list_recent_files / list_folder."
    lines = []
    for h in hits:
        parent = (h.get("parentReference") or {}).get("path", "")
        lines.append(
            f"- {h.get('name')}  (id={h['id']})  path={parent}/{h.get('name')}  "
            f"size={h.get('size')}  modified={h.get('lastModifiedDateTime')}"
        )
    return f"Found {len(lines)} item(s):\n" + "\n".join(lines)


@mcp.tool()
async def read_file_by_name(filename: str) -> str:
    """Find a file by name only (no path) and return its content (Office/text/images)."""
    auth.require_user()
    item = graph.find_file_by_name(filename)
    if not item:
        return (
            f"No file matching '{filename}'. "
            "Try search_onedrive with a shorter query or the stem without extension, "
            "or list_recent_files."
        )
    header = (
        f"File: {item.get('name')}\n"
        f"id: {item['id']}\n"
        f"path: {(item.get('parentReference') or {}).get('path', '')}/{item.get('name')}\n"
        f"size: {item.get('size')}  modified: {item.get('lastModifiedDateTime')}\n"
        f"---\n"
    )
    return header + graph.read_office_or_text(item)


@mcp.tool()
async def read_file_by_id(item_id: str) -> str:
    """Read content of a OneDrive file by Graph item id (from search_onedrive)."""
    auth.require_user()
    meta = graph.item_metadata(item_id)
    header = f"File: {meta.get('name')}\nid: {meta['id']}\n---\n"
    return header + graph.read_office_or_text(meta)


@mcp.tool()
async def list_folder(path: str = "/", top: int = 100) -> str:
    """List files/folders in a OneDrive path (default root '/')."""
    auth.require_user()
    items = graph.list_folder(path, top=top)
    if not items:
        return "Empty folder."
    lines = []
    for it in items:
        kind = "folder" if "folder" in it else "file"
        lines.append(f"- [{kind}] {it.get('name')}  (id={it['id']})  size={it.get('size')}")
    return "\n".join(lines)


@mcp.tool()
async def list_recent_files(top: int = 30) -> str:
    """List recently used OneDrive files."""
    auth.require_user()
    items = graph.list_recent_files(top=top)
    if not items:
        return "No recent files."
    return "\n".join(
        f"- {it.get('name')}  (id={it['id']})  modified={it.get('lastModifiedDateTime')}"
        for it in items
    )


@mcp.tool()
async def upload_new_file_to_appfolder(
    filename: str,
    content: str,
    content_encoding: str = "text",
) -> str:
    """
    Create a NEW file in the app OneDrive AppFolder only (never overwrites).

    content_encoding:
      - "text" (default): for .docx/.xlsx/.pptx the server builds a real styled
        Office file from your text. Other extensions are stored as UTF-8 text.
          docx: blank line = new paragraph; lines starting with - or * = bullets
          xlsx: one row per line; tabs or commas = columns (first row = header)
          pptx: slides separated by a line containing only --- ;
                first slide is a dark title slide; body lines starting with -
                become bullets
      - "base64": content is base64 of a finished file; uploaded as-is
        (use this when you already have binary bytes for any allowed type).
    """
    auth.require_user()
    try:
        meta = graph.upload_new_file_to_appfolder(
            filename, content, content_encoding=content_encoding
        )
    except (ValueError, FileExistsError, PermissionError) as e:
        return f"Upload refused: {e}"
    return (
        f"Created new file in AppFolder: {meta.get('name')} "
        f"(id={meta.get('id')}) size={meta.get('size')}"
    )


# Backwards-compatible name
@mcp.tool()
async def upload_new_file_to_ai(filename: str, content: str) -> str:
    """Alias for upload_new_file_to_appfolder (AppFolder, not /AI)."""
    return await upload_new_file_to_appfolder(filename, content)


# ---- OneNote ----

@mcp.tool()
async def list_onenote_notebooks() -> str:
    """List all OneNote notebooks."""
    auth.require_user()
    nbs = graph.list_notebooks()
    if not nbs:
        return "No notebooks."
    return "\n".join(f"- {nb.get('displayName')}  (id={nb['id']})" for nb in nbs)


@mcp.tool()
async def list_onenote_sections(notebook_id: str | None = None) -> str:
    """List sections. Pass notebook_id to limit to one notebook, or omit for all."""
    auth.require_user()
    secs = graph.list_sections(notebook_id)
    if not secs:
        return "No sections."
    return "\n".join(
        f"- {s.get('displayName')}  (id={s['id']})  "
        f"notebook={s.get('parentNotebook', {}).get('displayName', '?')}"
        for s in secs
    )


@mcp.tool()
async def list_onenote_pages(section_id: str) -> str:
    """List pages inside a section."""
    auth.require_user()
    pages = graph.list_pages(section_id)
    if not pages:
        return "No pages."
    return "\n".join(
        f"- {p.get('title')}  (id={p['id']})  created={p.get('createdDateTime')}"
        for p in pages
    )


@mcp.tool()
async def read_onenote_page(page_id: str, inline_images: bool = True) -> str:
    """Get full HTML of a OneNote page; images/drawings inlined as data URIs by default."""
    auth.require_user()
    return graph.get_page_html(page_id, inline_images=inline_images)


@mcp.tool()
async def find_and_read_onenote_page(title: str) -> str:
    """Find a OneNote page by title (partial match) and return HTML with images inlined."""
    auth.require_user()
    page = graph.find_onenote_page_by_title(title)
    if not page:
        return f"No OneNote page whose title contains '{title}'."
    html = graph.get_page_html(page["id"], inline_images=True)
    return f"Title: {page.get('title')}\nid: {page['id']}\n\n{html}"


# ---- To Do ----

@mcp.tool()
async def list_todo_lists() -> str:
    """List Microsoft To Do lists."""
    auth.require_user()
    lists = graph.list_todo_lists()
    if not lists:
        return "No To Do lists."
    return "\n".join(f"- {lst.get('displayName')}  (id={lst['id']})" for lst in lists)


@mcp.tool()
async def list_todo_tasks(list_id: str) -> str:
    """List tasks in a To Do list."""
    auth.require_user()
    tasks = graph.list_tasks(list_id)
    if not tasks:
        return "No tasks."
    lines = []
    for t in tasks:
        status = t.get("status", "?")
        due = (t.get("dueDateTime") or {}).get("dateTime", "")
        lines.append(f"- [{status}] {t.get('title')}  (id={t['id']})  due={due}")
    return "\n".join(lines)


@mcp.tool()
async def create_todo_task(
    list_id: str, title: str, body: str | None = None, due_iso: str | None = None
) -> str:
    """Create a NEW To Do task only. Cannot complete, update, or delete existing tasks."""
    auth.require_user()
    try:
        t = graph.create_task(list_id, title, body=body, due_iso=due_iso)
    except PermissionError as e:
        return f"Refused: {e}"
    return f"Created task: {t.get('title')} (id={t.get('id')})"


# ---- Mail (list = basic; full body requires sender PIN) ----

@mcp.tool()
async def list_emails(folder: str = "inbox", top: int = 15, search: str | None = None) -> str:
    """List emails (metadata only). Full body is not included; use read_email with pin."""
    auth.require_user()
    msgs = graph.list_messages(folder=folder, top=top, search=search)
    if not msgs:
        return "No messages."
    lines = []
    for m in msgs:
        frm = ((m.get("from") or {}).get("emailAddress") or {}).get("address", "?")
        read = "read" if m.get("isRead") else "UNREAD"
        lines.append(
            f"- [{read}] {m.get('subject')}  from={frm}  "
            f"at={m.get('receivedDateTime')}  (id={m['id']})\n"
            f"  preview: {(m.get('bodyPreview') or '')[:160]}"
        )
    return "\n".join(lines)


@mcp.tool()
async def read_email(message_id: str, pin: str) -> str:
    """Read full email body by message id. Requires a valid pin for that message."""
    auth.require_user()
    try:
        m = graph.get_message_with_pin(message_id, pin)
    except PermissionError:
        return "Refused: invalid or missing pin."
    frm = ((m.get("from") or {}).get("emailAddress") or {}).get("address", "?")
    body = (m.get("body") or {}).get("content", "")
    return (
        f"Subject: {m.get('subject')}\nFrom: {frm}\n"
        f"Received: {m.get('receivedDateTime')}\n---\n{body}"
    )


# ---- Calendar (create allowed; no edit/delete) ----

@mcp.tool()
async def list_calendar_events(days_ahead: int = 7, top: int = 30) -> str:
    """List upcoming calendar events."""
    auth.require_user()
    events = graph.list_events(days_ahead=days_ahead, top=top)
    if not events:
        return "No upcoming events."
    lines = []
    for e in events:
        start = (e.get("start") or {}).get("dateTime", "?")
        end = (e.get("end") or {}).get("dateTime", "?")
        loc = (e.get("location") or {}).get("displayName", "")
        lines.append(
            f"- {e.get('subject')}  {start} → {end}  loc={loc}  (id={e['id']})"
        )
    return "\n".join(lines)


@mcp.tool()
async def create_calendar_event(
    subject: str,
    start_iso: str,
    end_iso: str,
    body: str | None = None,
    location: str | None = None,
    timezone: str = "UTC",
) -> str:
    """
    Create a NEW calendar event only. Cannot edit or delete existing events.
    Times as ISO local date-time; timezone default UTC.
    Example start_iso: 2026-11-10T17:00:00
    """
    auth.require_user()
    try:
        e = graph.create_event(
            subject, start_iso, end_iso, body=body, location=location, timezone=timezone
        )
    except PermissionError as e:
        return f"Refused: {e}"
    return f"Created event: {e.get('subject')} (id={e.get('id')})"


# ---- Contacts ----

@mcp.tool()
async def list_contacts(top: int = 50) -> str:
    """List Outlook contacts."""
    auth.require_user()
    contacts = graph.list_contacts(top=top)
    if not contacts:
        return "No contacts."
    lines = []
    for c in contacts:
        emails = ", ".join(
            (e.get("address") or "") for e in (c.get("emailAddresses") or [])
        )
        lines.append(
            f"- {c.get('displayName')}  email={emails}  "
            f"phone={c.get('mobilePhone') or ''}  (id={c['id']})"
        )
    return "\n".join(lines)


@mcp.tool()
async def search_contacts(query: str) -> str:
    """Search contacts by display name prefix."""
    auth.require_user()
    contacts = graph.search_contacts(query)
    if not contacts:
        return "No contacts matched."
    lines = []
    for c in contacts:
        emails = ", ".join(
            (e.get("address") or "") for e in (c.get("emailAddresses") or [])
        )
        lines.append(f"- {c.get('displayName')}  email={emails}  (id={c['id']})")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# HTTP: login pages + pure ASGI auth middleware
# ---------------------------------------------------------------------------

LOGIN_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1"/>
  <title>Microsoft MCP – Sign in</title>
  <link rel="icon" href="/favicon.ico" type="image/x-icon"/>
  <style>
    body {{ font-family: system-ui, sans-serif; max-width: 520px; margin: 4rem auto; padding: 0 1rem; color: #111; }}
    h1 {{ font-size: 1.4rem; }}
    a.btn {{ display: inline-block; background: #2f2f2f; color: #fff; padding: 0.75rem 1.25rem;
             border-radius: 8px; text-decoration: none; margin-top: 1rem; }}
    a.btn:hover {{ background: #000; }}
    .muted {{ color: #666; font-size: 0.95rem; }}
  </style>
</head>
<body>
  <h1>Microsoft MCP</h1>
  <p class="muted">Sign in with your personal Microsoft account. You will get a session token to put in your MCP client as <code>Authorization: Bearer …</code>.</p>
  <a class="btn" href="{auth_url}">Sign in with Microsoft</a>
</body>
</html>
"""

SUCCESS_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1"/>
  <title>Microsoft MCP – Connected</title>
  <link rel="icon" href="/favicon.ico" type="image/x-icon"/>
  <style>
    body {{ font-family: system-ui, sans-serif; max-width: 640px; margin: 3rem auto; padding: 0 1rem; color: #111; }}
    code, pre {{ background: #f4f4f4; border-radius: 6px; padding: 0.2rem 0.4rem; font-size: 0.9rem; }}
    pre {{ padding: 1rem; overflow-x: auto; word-break: break-all; white-space: pre-wrap; }}
    .ok {{ color: #0a7; font-weight: 600; }}
    .box {{ border: 1px solid #ddd; border-radius: 10px; padding: 1rem 1.25rem; margin: 1.25rem 0; }}
  </style>
</head>
<body>
  <p class="ok">Signed in as {display_name} {email}</p>
  <h1>Connect this account in Grok / Claude</h1>
  <div class="box">
    <p><strong>MCP URL</strong></p>
    <pre>{mcp_url}</pre>
    <p><strong>Authorization header</strong> (required)</p>
    <pre>Authorization: Bearer {session_token}</pre>
  </div>
  <p>Without this Bearer token, tools cannot access any Microsoft data.</p>
</body>
</html>
"""

ERROR_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8"/>
  <title>Login error</title>
  <link rel="icon" href="/favicon.ico" type="image/x-icon"/>
  <style>body{{font-family:system-ui,sans-serif;max-width:560px;margin:3rem auto;padding:0 1rem}}</style>
</head>
<body>
  <h1>Login failed</h1>
  <pre>{message}</pre>
  <p><a href="/login">Try again</a></p>
</body>
</html>
"""


async def homepage(request: Request) -> Response:
    return HTMLResponse(
        "<!DOCTYPE html><html><head>"
        "<link rel='icon' href='/favicon.ico' type='image/x-icon'/>"
        "<title>Microsoft MCP</title></head>"
        "<body style='font-family:system-ui;margin:3rem'>"
        "<h1>Microsoft MCP</h1>"
        "<p><a href='/login'>Sign in with Microsoft</a> to get your personal session token.</p>"
        "<p>MCP endpoint: <code>/mcp</code></p><p>Auth: OAuth (ChatGPT/Grok) via <code>/authorize</code>, or manual Bearer from <a href='/login'>/login</a>.</p>"
        "</body></html>"
    )


async def favicon(request: Request) -> Response:
    """Serve /favicon.ico so clients (and Claude) don't fall back to the ngrok icon."""
    from pathlib import Path

    path = Path(__file__).resolve().parent / "favicon.ico"
    if not path.is_file():
        return Response(status_code=404)
    data = path.read_bytes()
    return Response(content=data, media_type="image/x-icon")


async def login_page(request: Request) -> Response:
    try:
        auth_url = auth.start_login()
    except Exception as e:
        return HTMLResponse(ERROR_HTML.format(message=str(e)), status_code=500)
    return HTMLResponse(LOGIN_HTML.format(auth_url=auth_url))


async def auth_callback(request: Request) -> Response:
    """Microsoft OAuth callback — also completes MCP client OAuth when applicable."""
    from starlette.responses import RedirectResponse

    try:
        params = dict(request.query_params)
        if "error" in params:
            raise RuntimeError(
                params.get("error_description") or params.get("error") or "OAuth error"
            )
        result = auth.complete_oauth_after_microsoft(params)
    except Exception as e:
        return HTMLResponse(ERROR_HTML.format(message=str(e)), status_code=400)

    if result.get("mode") == "oauth_redirect":
        return RedirectResponse(result["redirect_to"], status_code=302)

    email = f"&lt;{result['email']}&gt;" if result.get("email") else ""
    mcp_url = auth.PUBLIC_BASE_URL.rstrip("/") + "/mcp"
    ttl_days = max(1, auth.SESSION_TTL // 86400)
    return HTMLResponse(
        SUCCESS_HTML.format(
            display_name=result["display_name"],
            email=email,
            mcp_url=mcp_url,
            session_token=result["session_token"],
            ttl_days=ttl_days,
        )
    )


# ---- MCP OAuth 2.1 endpoints (ChatGPT / Grok / clients without custom headers) ----

async def well_known_protected_resource(request: Request) -> Response:
    return JSONResponse(auth.oauth_protected_resource_metadata())


async def well_known_authorization_server(request: Request) -> Response:
    return JSONResponse(auth.oauth_authorization_server_metadata())


async def oauth_register(request: Request) -> Response:
    try:
        body = await request.json()
    except Exception:
        body = {}
    try:
        out = auth.register_oauth_client(body if isinstance(body, dict) else {})
        return JSONResponse(out, status_code=201)
    except ValueError as e:
        return JSONResponse({"error": "invalid_client_metadata", "error_description": str(e)}, status_code=400)


async def oauth_authorize(request: Request) -> Response:
    """
    Authorization endpoint for MCP clients.
    Redirects the user to Microsoft sign-in, then back to the client with a code.
    """
    from starlette.responses import RedirectResponse

    q = request.query_params
    client_id = q.get("client_id") or ""
    redirect_uri = q.get("redirect_uri") or ""
    state = q.get("state") or ""
    response_type = q.get("response_type") or "code"
    code_challenge = q.get("code_challenge")
    code_challenge_method = q.get("code_challenge_method")
    scope = q.get("scope")

    if response_type != "code":
        return JSONResponse(
            {"error": "unsupported_response_type"}, status_code=400
        )
    if not client_id or not redirect_uri:
        return JSONResponse(
            {"error": "invalid_request", "error_description": "client_id and redirect_uri required"},
            status_code=400,
        )
    try:
        ms_url = auth.start_oauth_authorize(
            client_id=client_id,
            redirect_uri=redirect_uri,
            state=state,
            code_challenge=code_challenge,
            code_challenge_method=code_challenge_method,
            scope=scope,
        )
    except ValueError as e:
        return JSONResponse(
            {"error": "invalid_request", "error_description": str(e)}, status_code=400
        )
    except Exception as e:
        return HTMLResponse(ERROR_HTML.format(message=str(e)), status_code=500)
    return RedirectResponse(ms_url, status_code=302)


async def oauth_token(request: Request) -> Response:
    """Token endpoint — exchanges authorization code (or refresh_token) for Bearer access token."""
    content_type = (request.headers.get("content-type") or "").lower()
    form: dict = {}
    if "application/json" in content_type:
        try:
            form = await request.json()
        except Exception:
            form = {}
    else:
        try:
            form = dict(await request.form())
        except Exception:
            form = {}

    # client_secret_basic
    client_id = form.get("client_id")
    client_secret = form.get("client_secret")
    auth_header = request.headers.get("authorization") or ""
    if auth_header.lower().startswith("basic "):
        import base64

        try:
            raw = base64.b64decode(auth_header[6:].strip()).decode("utf-8")
            cid, _, csec = raw.partition(":")
            client_id = client_id or cid
            client_secret = client_secret or csec
        except Exception:
            pass

    grant_type = form.get("grant_type") or ""
    try:
        token = auth.exchange_oauth_token(
            grant_type=grant_type,
            code=form.get("code"),
            redirect_uri=form.get("redirect_uri"),
            client_id=client_id,
            client_secret=client_secret,
            code_verifier=form.get("code_verifier"),
            refresh_token=form.get("refresh_token"),
        )
        return JSONResponse(token)
    except PermissionError as e:
        msg = str(e)
        err = "invalid_grant"
        if "invalid_client" in msg:
            err = "invalid_client"
        return JSONResponse(
            {"error": err, "error_description": msg}, status_code=400
        )
    except ValueError as e:
        return JSONResponse(
            {"error": "invalid_request", "error_description": str(e)}, status_code=400
        )


class SessionAuthASGI:
    """
    Fail-closed ASGI auth.

    Public paths (no Bearer): login, OAuth discovery, authorize/token/register, static.
    EVERY other path — including all MCP protocol traffic — requires a valid
    Authorization: Bearer token that maps to exactly one Microsoft user.
    No session-id fallback. No anonymous Graph access. Ever.
    """

    # Exact paths that may be hit without a Bearer token.
    PUBLIC_EXACT = frozenset({
        "/",
        "/login",
        "/auth/callback",
        "/favicon.ico",
        "/favicon.png",
        "/.well-known/oauth-protected-resource",
        "/.well-known/oauth-authorization-server",
        "/.well-known/openid-configuration",
        "/register",
        "/authorize",
        "/token",
    })

    def __init__(self, app: ASGIApp):
        self.app = app

    async def __call__(self, scope: Scope, receive: Receive, send: Send) -> None:
        if scope["type"] != "http":
            await self.app(scope, receive, send)
            return

        # Always wipe identity first — never inherit from another request/task.
        auth.clear_request_user()

        raw_path = scope.get("path") or ""
        # Normalize: strip trailing slash except root
        path = raw_path if raw_path == "/" else raw_path.rstrip("/")

        headers = {
            k.decode("latin-1").lower(): v.decode("latin-1")
            for k, v in (scope.get("headers") or [])
        }

        def _unauthorized(message: str) -> JSONResponse:
            meta = auth.public_base() + "/.well-known/oauth-protected-resource"
            return JSONResponse(
                {"error": "Unauthorized", "message": message},
                status_code=401,
                headers={
                    "www-authenticate": (
                        f'Bearer realm="mcp", resource_metadata="{meta}"'
                    ),
                    "cache-control": "no-store",
                },
            )

        # Public endpoints: no Bearer required
        if path in self.PUBLIC_EXACT:
            try:
                await self.app(scope, receive, send)
            finally:
                auth.clear_request_user()
            return

        # ---- Everything else is protected (MCP and any unknown route) ----
        auth_header = headers.get("authorization") or ""
        token = ""
        if auth_header.lower().startswith("bearer "):
            token = auth_header[7:].strip()

        if not token:
            print(
                f"[auth] 401 no bearer path={path!r} "
                f"client={headers.get('x-forwarded-for') or headers.get('x-real-ip') or '?'}",
                flush=True,
            )
            body = _unauthorized(
                "Authentication required. Complete Microsoft sign-in via OAuth "
                "before using this server. No token means no data."
            )
            await body(scope, receive, send)
            return

        try:
            auth.set_request_user(token, mcp_session_id=None)
        except PermissionError as e:
            print(f"[auth] 401 invalid token path={path!r}: {e}", flush=True)
            body = _unauthorized(str(e))
            await body(scope, receive, send)
            return
        except Exception as e:
            print(f"[auth] 401 auth error path={path!r}: {e}", flush=True)
            body = _unauthorized(f"Auth error: {e}")
            await body(scope, receive, send)
            return

        bound = auth.current_user_id.get()
        session = auth.resolve_session(token)
        if not session or not bound or str(session.get("user_id")) != str(bound):
            auth.clear_request_user()
            print(f"[auth] 401 user mismatch path={path!r}", flush=True)
            body = _unauthorized("Session/user mismatch — refused.")
            await body(scope, receive, send)
            return

        print(f"[auth] ok user={bound} path={path!r}", flush=True)
        try:
            await self.app(scope, receive, send)
        finally:
            auth.clear_request_user()



def build_app() -> Starlette:
    import contextlib

    @contextlib.asynccontextmanager
    async def lifespan(app: Starlette):
        async with mcp.session_manager.run():
            yield

    public = os.environ.get("MS_PUBLIC_BASE_URL", "http://localhost:8000")
    public_host = public.split("://", 1)[-1].split("/")[0]
    security = TransportSecuritySettings(
        enable_dns_rebinding_protection=True,
        allowed_hosts=[
            "127.0.0.1:*",
            "localhost:*",
            "[::1]:*",
            public_host,
            f"{public_host}:*",
        ],
        allowed_origins=[
            "http://localhost:*",
            "http://127.0.0.1:*",
            f"https://{public_host}",
            f"http://{public_host}",
        ],
    )

    mcp_app = mcp.streamable_http_app(transport_security=security)

    routes = [
        Route("/", homepage),
        Route("/login", login_page),
        Route("/auth/callback", auth_callback),
        Route("/favicon.ico", favicon),
        # MCP OAuth 2.1 discovery + endpoints (ChatGPT / Grok)
        Route("/.well-known/oauth-protected-resource", well_known_protected_resource),
        Route("/.well-known/oauth-authorization-server", well_known_authorization_server),
        Route("/.well-known/openid-configuration", well_known_authorization_server),
        Route("/register", oauth_register, methods=["POST"]),
        Route("/authorize", oauth_authorize, methods=["GET"]),
        Route("/token", oauth_token, methods=["POST"]),
        Mount("/", app=mcp_app),
    ]
    app = Starlette(routes=routes, lifespan=lifespan)
    # Pure ASGI wrapper outermost
    return SessionAuthASGI(app)  # type: ignore[return-value]


# ---------------------------------------------------------------------------
# Entry
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    transport = os.environ.get("MCP_TRANSPORT", "http").lower()
    host = os.environ.get("MCP_HOST", "0.0.0.0")
    port = int(os.environ.get("MCP_PORT", "8000"))

    if transport in ("http", "streamable-http", "sse"):
        import uvicorn

        print(f"Microsoft MCP (multi-user + OAuth) on http://{host}:{port}", flush=True)
        print(f"  Login page:   http://{host}:{port}/login", flush=True)
        print(f"  MCP endpoint: http://{host}:{port}/mcp", flush=True)
        print(f"  OAuth AS:     http://{host}:{port}/.well-known/oauth-authorization-server", flush=True)
        print(f"  MS redirect:  {auth.REDIRECT_URI}", flush=True)
        print(f"  DATA_DIR:     {auth.DATA_DIR}", flush=True)
        print("  Auth mode:    FAIL-CLOSED — every non-public path requires Bearer", flush=True)
        uvicorn.run(build_app(), host=host, port=port)
    else:
        print(
            "stdio transport is not supported in multi-user mode.\n"
            "Run with MCP_TRANSPORT=http and open /login in a browser.",
            flush=True,
        )
        raise SystemExit(1)
