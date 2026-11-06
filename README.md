# Microsoft Graph MCP (multi-user, restricted write)

MCP server for **personal Microsoft accounts**. Each person signs in with Microsoft in the browser, gets their own session token, and the AI only ever sees **that** user’s data.

## Permission model

| Action | Allowed? |
|--------|----------|
| Read OneDrive everywhere (`Files.Read` / `Files.Read.All`) | Yes |
| Create **new** files in app **AppFolder** only (`Files.ReadWrite.AppFolder`) | Yes (common types only) |
| Overwrite / update existing OneDrive files | **No** |
| Create / edit / delete OneNote | **No** (read-only) |
| Create **new** To Do tasks | Yes |
| Complete / update / delete To Do tasks or lists | **No** |
| List emails (subject/from/preview — Mail.ReadBasic fields) | Yes |
| Read full email body | Only with correct **sender PIN** |
| Send email | **No** |
| Create calendar events | Yes |
| Edit / delete calendar events | **No** |
| Delete anything | **No** |

### Email PIN
Full body requires `read_email(message_id, pin)`.
PIN = lengths of each segment of the sender address when split on `.` and `@`, joined by `_`.
Example: `firstname.lastname@outlook.com` → `9_8_7_3`.

### AppFolder writes
Uploads use Graph ` /me/drive/special/approot ` (application folder). The app cannot write outside that folder even with a broader token.

## How access control works

1. User opens `https://your-server/login`
2. Microsoft login page → consent to the restricted scopes
3. Callback creates a **session token** shown on the success page
4. User configures Grok/Claude MCP with:
   - **URL:** `https://your-server/mcp`
   - **Header:** `Authorization: Bearer <session_token>`
5. Every tool call is bound to that session’s Microsoft account. No token → 401. Wrong user → their own data only.

Tokens live only on your machine (`./data/`). The public URL without a valid Bearer cannot read anything.

---

## 1. Entra app registration (one-time)

1. [Entra app registrations](https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade) → **New registration**
2. **Supported account types:** *Personal Microsoft accounts only* (or “any org + personal”)
3. **Authentication** → **Add a platform** → **Web**
   - Redirect URI: `http://localhost:8000/auth/callback`
   - Also add your public tunnel, e.g. `https://xxxx.ngrok-free.app/auth/callback`
4. Under **Implicit grant** leave unchecked (we use auth code + PKCE)
5. **Advanced settings** → **Allow public client flows** = **Yes** (needed for PKCE public client)
6. **API permissions** → Microsoft Graph → **Delegated**:
   - `User.Read`
   - `Files.Read` and/or `Files.Read.All` (read everywhere)
   - `Files.ReadWrite.AppFolder` (write only into app folder)
   - `Notes.Read`
   - `Tasks.ReadWrite` (create tasks only; update/delete blocked in code)
   - `Mail.ReadBasic` (list/search without body)
   - `Mail.Read` (full body; still gated by sender PIN in code)
   - `Calendars.ReadWrite` (create events only; edit/delete blocked in code)
   - `Contacts.Read`
7. Copy **Application (client) ID**

No client secret required when using public client + PKCE (default).

**Important:** after changing scopes, users must sign in again and re-consent.

---

## 2. Install

```powershell
cd microsoft-mcp-server
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
```

---

## 3. Environment variables

```powershell
$env:MS_CLIENT_ID = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

# Must match a Redirect URI registered in Entra exactly
$env:MS_REDIRECT_URI = "http://localhost:8000/auth/callback"
# When using a tunnel, set BOTH to the public URLs:
# $env:MS_REDIRECT_URI = "https://xxxx.ngrok-free.app/auth/callback"
# $env:MS_PUBLIC_BASE_URL = "https://xxxx.ngrok-free.app"

$env:MS_PUBLIC_BASE_URL = "http://localhost:8000"
$env:MCP_TRANSPORT = "http"
$env:MCP_PORT = "8000"
$env:MCP_HOST = "0.0.0.0"
```

Optional:

| Variable | Meaning |
|----------|---------|
| `MS_CLIENT_SECRET` | If set, uses confidential client instead of public+PKCE |
| `MS_SESSION_TTL` | Session lifetime in seconds (default 30 days) |
| `MS_DATA_DIR` | Where sessions/tokens are stored (default `./data`) |

---

## 4. Run

```powershell
python server.py
```

Open **http://localhost:8000/login** (or your tunnel URL + `/login`).

After Microsoft sign-in you’ll see:

- MCP URL  
- `Authorization: Bearer …` token  

Paste those into Grok’s custom MCP / connector settings (URL + header).

---

## 5. Tunnel (Grok web / mobile)

Grok cannot reach `localhost`. Expose the server:

**ngrok example**

```powershell
ngrok http 8000
```

Then set:

```powershell
$env:MS_REDIRECT_URI = "https://YOUR-SUBDOMAIN.ngrok-free.app/auth/callback"
$env:MS_PUBLIC_BASE_URL = "https://YOUR-SUBDOMAIN.ngrok-free.app"
```

Register that same redirect URI in Entra, restart `server.py`, open `https://YOUR-SUBDOMAIN.ngrok-free.app/login`.

**Important:** even with a public tunnel, calls to `/mcp` without a valid Bearer session token get **401**. Each user must complete `/login` themselves.

---

## Tools

### Read
| Tool | Purpose |
|------|---------|
| `whoami` | Which Microsoft account is active |
| `search_onedrive` | Find files by name/content |
| `read_file_by_name` | Open file by **name only** |
| `read_file_by_id` | Open by Graph id |
| `list_folder` / `list_recent_files` | Browse OneDrive |
| `list_onenote_notebooks` / `sections` / `pages` | Browse OneNote |
| `read_onenote_page` / `find_and_read_onenote_page` | Read OneNote (images inlined) |
| `list_todo_lists` / `list_todo_tasks` | To Do (read) |
| `list_emails` / `read_email` | Mail (read) |
| `list_calendar_events` | Calendar (read) |
| `list_contacts` / `search_contacts` | Contacts (read) |

### Write (strictly limited)
| Tool | Purpose |
|------|---------|
| `upload_new_file_to_appfolder` | Create a **new** file in AppFolder only; refuses overwrite and disallowed extensions |
| `create_todo_task` | Create a **new** task in an existing list; cannot complete/update/delete |
| `create_calendar_event` | Create a **new** calendar event; cannot edit/delete |

### Mail
| Tool | Purpose |
|------|---------|
| `list_emails` | Basic fields only (no full body) |
| `read_email(message_id, pin)` | Full body **only** if sender PIN matches |

---

## Security notes

- Scopes are intentionally narrow; extra guards in code refuse overwrite, non-`/AI` paths, OneNote writes, mail send, deletes, and task updates.
- Session tokens are random 32-byte values; treat them like passwords.
- Microsoft refresh tokens stay on disk under `./data/users/` — do not commit that folder.
- Different session tokens map to different Microsoft users; there is no cross-account access.

---

## Troubleshooting

**AADSTS70002 / device flow errors**  
This build no longer uses device code. Use `/login` in a browser.

**Redirect URI mismatch**  
`MS_REDIRECT_URI` must match Entra **exactly** (scheme, host, path, no trailing slash mismatch).

**401 on tool calls**  
MCP client is not sending `Authorization: Bearer <token>`. Re-open `/login` and copy the new token into the connector headers.

**Consent / scope errors after update**  
Users must visit `/login` again and accept the new permission set.

**`streamable_http_app` missing**  
Upgrade: `pip install -U "mcp[cli]>=2.0"`.
