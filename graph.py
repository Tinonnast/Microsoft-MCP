"""Microsoft Graph helpers used by the MCP tools (read + write)."""
from __future__ import annotations

import base64
import io
import re
from typing import Any
from urllib.parse import quote

import requests
from bs4 import BeautifulSoup

from auth import graph_headers, current_user_id, current_access_token

GRAPH = "https://graph.microsoft.com/v1.0"


def _assert_user_bound() -> None:
    """Hard stop: no Graph traffic unless this request has a bound user + token."""
    if not current_user_id.get() or not current_access_token.get():
        raise RuntimeError(
            "Refused Graph call: no authenticated user on this request."
        )


# ---------------------------------------------------------------------------
# HTTP
# ---------------------------------------------------------------------------

def gget(path: str, params: dict | None = None) -> Any:
    _assert_user_bound()
    r = requests.get(f"{GRAPH}{path}", headers=graph_headers(), params=params, timeout=60)
    r.raise_for_status()
    if r.headers.get("content-type", "").startswith("application/json"):
        return r.json()
    return r.content


def gget_all(path: str, params: dict | None = None, max_pages: int = 20) -> list:
    """Follow @odata.nextLink up to max_pages."""
    params = dict(params or {})
    items: list = []
    data = gget(path, params=params)
    items.extend(data.get("value", []))
    pages = 1
    while data.get("@odata.nextLink") and pages < max_pages:
        r = requests.get(
            data["@odata.nextLink"], headers=graph_headers(), timeout=60
        )
        r.raise_for_status()
        data = r.json()
        items.extend(data.get("value", []))
        pages += 1
    return items


def gpost(
    path: str,
    json: dict | None = None,
    data: bytes | None = None,
    headers: dict | None = None,
) -> Any:
    _assert_user_bound()
    h = graph_headers()
    if headers:
        h.update(headers)
    r = requests.post(f"{GRAPH}{path}", headers=h, json=json, data=data, timeout=120)
    r.raise_for_status()
    if r.content and r.headers.get("content-type", "").startswith("application/json"):
        return r.json()
    return r.content


def gpatch(path: str, json: dict) -> Any:
    _assert_user_bound()
    r = requests.patch(f"{GRAPH}{path}", headers=graph_headers(), json=json, timeout=60)
    r.raise_for_status()
    return r.json() if r.content else None


def gdelete(path: str) -> None:
    _assert_user_bound()
    r = requests.delete(f"{GRAPH}{path}", headers=graph_headers(), timeout=60)
    r.raise_for_status()


def gput_bytes(path: str, content: bytes, content_type: str = "application/octet-stream") -> Any:
    _assert_user_bound()
    h = graph_headers()
    h["Content-Type"] = content_type
    r = requests.put(f"{GRAPH}{path}", headers=h, data=content, timeout=180)
    r.raise_for_status()
    if r.content and r.headers.get("content-type", "").startswith("application/json"):
        return r.json()
    return r.content


# ---------------------------------------------------------------------------
# OneDrive / files
# ---------------------------------------------------------------------------

def search_files(query: str, top: int = 50) -> list[dict]:
    """
    Search OneDrive by name/content with pagination.
    Graph drive search is imperfect — we pull multiple pages and de-dupe by id.
    """
    q = (query or "").strip()
    if not q:
        return []

    # Escape single quotes for OData string
    q_esc = q.replace("'", "''")
    path = f"/me/drive/root/search(q='{q_esc}')"
    # Request more than `top` then trim so we see past the first page of weak matches
    page_size = min(200, max(top, 50))
    items = gget_all(path, params={"$top": page_size}, max_pages=10)

    # De-dupe by id, preserve order
    seen: set[str] = set()
    out: list[dict] = []
    for it in items:
        iid = it.get("id")
        if iid and iid not in seen:
            seen.add(iid)
            out.append(it)
    return out[:top]


def find_file_by_name(name: str) -> dict | None:
    """
    Best-effort find by filename (no path needed).
    Tries exact match, then extension-aware search, then contains.
    """
    name = (name or "").strip()
    if not name:
        return None
    lower = name.lower()

    # 1) Search with full name
    hits = search_files(name, top=100)

    def score(h: dict) -> tuple:
        n = (h.get("name") or "").lower()
        exact = 0 if n == lower else 1
        start = 0 if n.startswith(lower) else 1
        # prefer newer
        mod = h.get("lastModifiedDateTime") or ""
        return (exact, start, 0 if lower in n else 1, mod)

    if hits:
        hits_sorted = sorted(hits, key=score)
        best = hits_sorted[0]
        if (best.get("name") or "").lower() == lower:
            return best
        if lower in (best.get("name") or "").lower():
            return best

    # 2) If name has extension, also search by stem and filter client-side
    if "." in name:
        stem, ext = name.rsplit(".", 1)
        more = search_files(stem, top=100)
        for h in more:
            n = (h.get("name") or "").lower()
            if n == lower:
                return h
        for h in more:
            n = (h.get("name") or "").lower()
            if n.endswith("." + ext.lower()) and stem.lower() in n:
                return h

    # 3) Search by extension only if query looks like "*.pptx" / "pptx"
    if lower.startswith("*.") or (len(lower) <= 5 and not "." in lower and len(lower) >= 3):
        ext = lower[2:] if lower.startswith("*.") else lower
        more = search_files(ext, top=200)
        more = [h for h in more if (h.get("name") or "").lower().endswith("." + ext)]
        if more:
            return sorted(more, key=lambda h: h.get("lastModifiedDateTime") or "", reverse=True)[0]

    return hits[0] if hits else None


def list_folder(path: str = "/", top: int = 100) -> list[dict]:
    """List children of a folder. path is like '/' or '/Documents'."""
    path = path.strip() or "/"
    if path == "/":
        items = gget_all("/me/drive/root/children", params={"$top": min(top, 200)})
    else:
        clean = path.strip("/")
        items = gget_all(
            f"/me/drive/root:/{quote(clean)}:/children",
            params={"$top": min(top, 200)},
        )
    return items[:top]


def list_recent_files(top: int = 30) -> list[dict]:
    data = gget("/me/drive/recent")
    return (data.get("value") or [])[:top]


def download_item_content(item_id: str) -> bytes:
    return gget(f"/me/drive/items/{item_id}/content")


def item_metadata(item_id: str) -> dict:
    return gget(f"/me/drive/items/{item_id}")


# Only these extensions may be uploaded as new files. No OneNote packages.
ALLOWED_UPLOAD_EXTENSIONS = {
    ".txt", ".md", ".csv", ".json", ".xml", ".html", ".htm",
    ".docx", ".doc", ".pptx", ".ppt", ".xlsx", ".xls", ".xlsm",
    ".pdf", ".rtf", ".odt", ".ods", ".odp",
    ".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp",
    ".py", ".js", ".ts", ".css", ".yaml", ".yml", ".toml", ".ini", ".log",
}


def _ensure_approot() -> dict:
    """
    Ensure the app's special folder exists (Files.ReadWrite.AppFolder).
    Graph creates it on first access for the registered application.
    """
    return gget("/me/drive/special/approot")


def _file_exists_in_approot(filename: str) -> bool:
    try:
        gget(f"/me/drive/special/approot:/{quote(filename)}")
        return True
    except requests.HTTPError as e:
        if e.response is not None and e.response.status_code == 404:
            return False
        raise





def _build_office_bytes(filename: str, text: str) -> tuple[bytes, str]:
    """Turn plain text into a real, presentable Office file for docx/xlsx/pptx."""
    name = (filename or "").lower()
    text = text if text is not None else ""

    if name.endswith(".docx"):
        from docx import Document
        from docx.shared import Pt, RGBColor, Inches
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        import io

        doc = Document()
        section = doc.sections[0]
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

        normalized = text.replace("\r\n", "\n").strip("\n")
        blocks = normalized.split("\n\n") if normalized else [""]
        first = True
        for block in blocks:
            lines = block.split("\n")
            if first and lines:
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                run = p.add_run(lines[0])
                run.bold = True
                run.font.size = Pt(22)
                run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
                for line in lines[1:]:
                    p2 = doc.add_paragraph(line)
                    for r in p2.runs:
                        r.font.size = Pt(11)
                first = False
            else:
                if lines and lines[0].lstrip().startswith(("- ", "* ", "• ")):
                    for line in lines:
                        t = line.lstrip()
                        for prefix in ("- ", "* ", "• "):
                            if t.startswith(prefix):
                                t = t[len(prefix):]
                                break
                        bp = doc.add_paragraph(t, style="List Bullet")
                        for r in bp.runs:
                            r.font.size = Pt(11)
                else:
                    p = doc.add_paragraph(lines[0] if lines else "")
                    for r in p.runs:
                        r.font.size = Pt(11)
                    for line in lines[1:]:
                        run = p.add_run()
                        run.add_break()
                        p.add_run(line)
        buf = io.BytesIO()
        doc.save(buf)
        return buf.getvalue(), (
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    if name.endswith((".xlsx", ".xlsm")):
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        import io
        import csv
        from io import StringIO

        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        raw = text.replace("\r\n", "\n").strip("\n")
        rows: list[list[str]] = []
        if "\t" in raw:
            for line in raw.split("\n"):
                rows.append(line.split("\t"))
        elif "," in raw and any(line.count(",") >= 1 for line in raw.split("\n")[:5]):
            rows = [list(r) for r in csv.reader(StringIO(raw))]
        else:
            for line in raw.split("\n"):
                rows.append([line])
        if not rows:
            rows = [[""]]

        header_fill = PatternFill("solid", fgColor="1F4E79")
        header_font = Font(bold=True, color="FFFFFF", size=11)
        cell_font = Font(size=11)
        thin = Border(
            left=Side(style="thin", color="D0D0D0"),
            right=Side(style="thin", color="D0D0D0"),
            top=Side(style="thin", color="D0D0D0"),
            bottom=Side(style="thin", color="D0D0D0"),
        )
        alt_fill = PatternFill("solid", fgColor="F2F7FB")

        for r_i, row in enumerate(rows, start=1):
            for c_i, val in enumerate(row, start=1):
                cell = ws.cell(row=r_i, column=c_i, value=val)
                cell.border = thin
                cell.alignment = Alignment(vertical="center", wrap_text=True)
                if r_i == 1 and len(rows) > 1:
                    cell.fill = header_fill
                    cell.font = header_font
                else:
                    cell.font = cell_font
                    if r_i % 2 == 0:
                        cell.fill = alt_fill

        for col in ws.columns:
            max_len = 0
            letter = col[0].column_letter
            for cell in col:
                max_len = max(max_len, len(str(cell.value or "")))
            ws.column_dimensions[letter].width = min(max(max_len + 2, 12), 40)

        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue(), (
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    if name.endswith(".pptx"):
        from pptx import Presentation
        from pptx.util import Inches, Pt, Emu
        from pptx.dml.color import RGBColor
        from pptx.enum.shapes import MSO_SHAPE
        from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
        import io

        # Widescreen 16:9
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

        ACCENT = RGBColor(0x1F, 0x4E, 0x79)      # deep blue
        ACCENT2 = RGBColor(0x2E, 0x75, 0xB6)     # lighter blue
        BG_DARK = RGBColor(0x1A, 0x1A, 0x2E)     # near-black
        WHITE = RGBColor(0xFF, 0xFF, 0xFF)
        MUTED = RGBColor(0x5B, 0x5B, 0x5B)
        LIGHT_BG = RGBColor(0xF5, 0xF7, 0xFA)

        chunks: list[str] = []
        cur: list[str] = []
        for line in text.replace("\r\n", "\n").split("\n"):
            if line.strip() in ("---", "===", "***"):
                chunks.append("\n".join(cur).strip())
                cur = []
            else:
                cur.append(line)
        chunks.append("\n".join(cur).strip())
        chunks = [c for c in chunks if c] or [""]

        def _add_accent_bar(slide, color=ACCENT):
            bar = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.12)
            )
            bar.fill.solid()
            bar.fill.fore_color.rgb = color
            bar.line.fill.background()

        def _add_footer(slide, page_num, total):
            box = slide.shapes.add_textbox(
                Inches(0.6), Inches(7.05), Inches(12), Inches(0.3)
            )
            tf = box.text_frame
            p = tf.paragraphs[0]
            p.text = f"{page_num} / {total}"
            p.font.size = Pt(10)
            p.font.color.rgb = MUTED
            p.alignment = PP_ALIGN.RIGHT

        total = len(chunks)
        for idx, chunk in enumerate(chunks):
            lines = [ln for ln in chunk.split("\n")]
            title = lines[0].strip() if lines else ""
            body_lines = [ln for ln in lines[1:] if ln.strip() != ""]

            # Title slide (first) gets a darker treatment
            if idx == 0 and total >= 1:
                slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
                # Full dark background shape
                bg = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE, Inches(0), Inches(0),
                    Inches(13.333), Inches(7.5)
                )
                bg.fill.solid()
                bg.fill.fore_color.rgb = BG_DARK
                bg.line.fill.background()

                # Accent stripe
                stripe = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE, Inches(0), Inches(2.8),
                    Inches(13.333), Inches(0.08)
                )
                stripe.fill.solid()
                stripe.fill.fore_color.rgb = ACCENT2
                stripe.line.fill.background()

                # Title
                tbox = slide.shapes.add_textbox(
                    Inches(0.8), Inches(2.0), Inches(11.7), Inches(1.0)
                )
                tf = tbox.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                p.text = title
                p.font.size = Pt(40)
                p.font.bold = True
                p.font.color.rgb = WHITE
                p.alignment = PP_ALIGN.CENTER

                # Subtitle = first body lines joined
                if body_lines:
                    sbox = slide.shapes.add_textbox(
                        Inches(1.5), Inches(3.2), Inches(10.3), Inches(1.5)
                    )
                    stf = sbox.text_frame
                    stf.word_wrap = True
                    sp = stf.paragraphs[0]
                    sp.text = " · ".join(body_lines[:3])
                    sp.font.size = Pt(18)
                    sp.font.color.rgb = RGBColor(0xB0, 0xB8, 0xC8)
                    sp.alignment = PP_ALIGN.CENTER
                continue

            # Content slides
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            # Light background
            bg = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(0), Inches(0),
                Inches(13.333), Inches(7.5)
            )
            bg.fill.solid()
            bg.fill.fore_color.rgb = LIGHT_BG
            bg.line.fill.background()

            _add_accent_bar(slide, ACCENT)

            # Left accent pill
            pill = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(0), Inches(0.12),
                Inches(0.15), Inches(7.38)
            )
            pill.fill.solid()
            pill.fill.fore_color.rgb = ACCENT2
            pill.line.fill.background()

            # Title
            tbox = slide.shapes.add_textbox(
                Inches(0.7), Inches(0.4), Inches(11.5), Inches(0.9)
            )
            tf = tbox.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = title
            p.font.size = Pt(28)
            p.font.bold = True
            p.font.color.rgb = ACCENT

            # Body
            if body_lines:
                bbox = slide.shapes.add_textbox(
                    Inches(0.7), Inches(1.5), Inches(11.8), Inches(5.2)
                )
                btf = bbox.text_frame
                btf.word_wrap = True
                first_p = True
                for bline in body_lines:
                    raw = bline.rstrip()
                    is_bullet = False
                    for prefix in ("- ", "* ", "• ", "– "):
                        if raw.lstrip().startswith(prefix):
                            raw = raw.lstrip()[len(prefix):]
                            is_bullet = True
                            break
                    if first_p:
                        bp = btf.paragraphs[0]
                        first_p = False
                    else:
                        bp = btf.add_paragraph()
                    if is_bullet:
                        bp.text = "•  " + raw
                        bp.level = 0
                    else:
                        bp.text = raw
                    bp.font.size = Pt(18)
                    bp.font.color.rgb = RGBColor(0x2D, 0x2D, 0x2D)
                    bp.space_after = Pt(10)

            _add_footer(slide, idx + 1, total)

        buf = io.BytesIO()
        prs.save(buf)
        return buf.getvalue(), (
            "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

    return text.encode("utf-8"), "text/plain; charset=utf-8"



def upload_new_file_to_appfolder(
    filename: str,
    content,
    content_type: str | None = None,
    content_encoding: str = "text",
) -> dict:
    """
    Upload a NEW file into the app OneDrive AppFolder only.
    Refuses overwrite.

    content_encoding:
      - "text" (default): build real OOXML for .docx/.xlsx/.pptx from plain text;
        other extensions stored as UTF-8 text.
      - "base64": decode content as base64 and upload the raw bytes as-is
        (AI can supply a pre-built binary file).
    """
    import base64

    filename = (filename or "").strip()
    if not filename or "/" in filename or "\\" in filename:
        raise ValueError("filename must be a simple name with no path separators")

    lower = filename.lower()
    ext = ""
    if "." in lower:
        ext = "." + lower.rsplit(".", 1)[-1]
    if ext not in ALLOWED_UPLOAD_EXTENSIONS:
        allowed = ", ".join(sorted(ALLOWED_UPLOAD_EXTENSIONS))
        raise ValueError(
            f"Extension '{ext or '(none)'}' is not allowed. "
            f"Allowed: {allowed}. OneNote and arbitrary binaries are blocked."
        )

    enc = (content_encoding or "text").strip().lower()
    if enc in ("base64", "b64"):
        if isinstance(content, bytes):
            raw = content
            # might already be decoded, or might be base64 bytes
            try:
                data = base64.b64decode(raw, validate=False)
                if len(data) == 0 and len(raw) > 0:
                    data = raw
            except Exception:
                data = raw
        else:
            s = (content or "").strip()
            # tolerate data URLs
            if "," in s and s.lower().startswith("data:"):
                s = s.split(",", 1)[1]
            try:
                data = base64.b64decode(s, validate=False)
            except Exception as e:
                raise ValueError(f"Invalid base64 content: {e}") from e
        if not data:
            raise ValueError("base64 content decoded to empty file")
        ctype = content_type or "application/octet-stream"
        if ext == ".docx":
            ctype = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        elif ext in (".xlsx", ".xlsm"):
            ctype = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        elif ext == ".pptx":
            ctype = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        elif ext == ".pdf":
            ctype = "application/pdf"
    else:
        # text path
        if isinstance(content, str):
            data, ctype = _build_office_bytes(filename, content)
        else:
            data = content
            ctype = content_type or "application/octet-stream"
            if ext in {".docx", ".xlsx", ".xlsm", ".pptx"} and data[:2] != b"PK":
                try:
                    text = data.decode("utf-8")
                except UnicodeDecodeError:
                    text = None
                if text is not None:
                    data, ctype = _build_office_bytes(filename, text)

    _ensure_approot()
    if _file_exists_in_approot(filename):
        raise FileExistsError(
            f"File '{filename}' already exists in the app folder. "
            "Updating or overwriting existing files is not permitted."
        )

    path = f"/me/drive/special/approot:/{quote(filename)}:/content"
    return gput_bytes(path, data, ctype)


# Backwards-compatible alias used by older tool names
upload_new_file_to_ai = upload_new_file_to_appfolder


def create_folder(parent_path: str, name: str) -> dict:
    raise PermissionError(
        "Creating folders is disabled. New files go only into the app's AppFolder."
    )


def delete_item(item_id: str) -> None:
    raise PermissionError("Deleting OneDrive items is disabled.")


def read_office_or_text(item: dict) -> str:
    name = item.get("name", "")
    mime = (item.get("file") or {}).get("mimeType", "")
    item_id = item["id"]
    raw = download_item_content(item_id)

    if mime.startswith("text/") or name.lower().endswith(
        (".txt", ".md", ".csv", ".json", ".xml", ".html", ".htm", ".py", ".js", ".ts", ".css")
    ):
        try:
            return raw.decode("utf-8")
        except UnicodeDecodeError:
            return raw.decode("latin-1", errors="replace")

    if mime.startswith("image/") or name.lower().endswith(
        (".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp")
    ):
        b64 = base64.b64encode(raw).decode("ascii")
        return f"[Image: {name}]\n![image](data:{mime or 'image/png'};base64,{b64})"

    if name.lower().endswith(".docx") or "wordprocessingml" in mime:
        try:
            from docx import Document

            doc = Document(io.BytesIO(raw))
            parts = []
            for p in doc.paragraphs:
                if p.text.strip():
                    parts.append(p.text)
            for t in doc.tables:
                rows = []
                for row in t.rows:
                    rows.append(" | ".join(c.text.strip() for c in row.cells))
                parts.append("\n".join(rows))
            return "\n\n".join(parts) or "[empty Word document]"
        except Exception as e:
            return f"[Could not parse Word file: {e}]"

    if name.lower().endswith((".xlsx", ".xlsm")) or "spreadsheetml" in mime:
        try:
            from openpyxl import load_workbook

            wb = load_workbook(io.BytesIO(raw), data_only=True, read_only=True)
            out = []
            for sheet in wb.worksheets:
                out.append(f"## Sheet: {sheet.title}")
                for row in sheet.iter_rows(values_only=True):
                    vals = [str(c) if c is not None else "" for c in row]
                    if any(vals):
                        out.append("\t".join(vals))
            return "\n".join(out) or "[empty Excel]"
        except Exception as e:
            return f"[Could not parse Excel: {e}]"

    if name.lower().endswith(".pptx") or "presentationml" in mime:
        try:
            from pptx import Presentation

            prs = Presentation(io.BytesIO(raw))
            out = []
            for i, slide in enumerate(prs.slides, 1):
                out.append(f"## Slide {i}")
                for shape in slide.shapes:
                    if hasattr(shape, "text") and shape.text.strip():
                        out.append(shape.text)
            return "\n\n".join(out) or "[empty PowerPoint]"
        except Exception as e:
            return f"[Could not parse PowerPoint: {e}]"

    b64_preview = base64.b64encode(raw[:2048]).decode("ascii")
    return (
        f"[Binary file: {name}]\n"
        f"mime={mime} size={len(raw)} bytes\n"
        f"base64_preview (first 2KB): {b64_preview}"
    )


# ---------------------------------------------------------------------------
# OneNote
# ---------------------------------------------------------------------------

def list_notebooks() -> list[dict]:
    return gget_all("/me/onenote/notebooks")


def list_sections(notebook_id: str | None = None) -> list[dict]:
    if notebook_id:
        return gget_all(f"/me/onenote/notebooks/{notebook_id}/sections")
    return gget_all("/me/onenote/sections")


def list_pages(section_id: str) -> list[dict]:
    return gget_all(f"/me/onenote/sections/{section_id}/pages")


def get_page_html(page_id: str, inline_images: bool = True) -> str:
    html = gget(f"/me/onenote/pages/{page_id}/content")
    if isinstance(html, bytes):
        html = html.decode("utf-8", errors="replace")

    if not inline_images:
        return html

    soup = BeautifulSoup(html, "lxml")
    for img in soup.find_all("img"):
        src = img.get("src") or img.get("data-fullres-src")
        if not src or not src.startswith("http"):
            continue
        try:
            r = requests.get(src, headers=graph_headers(), timeout=30)
            r.raise_for_status()
            mime = r.headers.get("content-type", "image/png").split(";")[0]
            b64 = base64.b64encode(r.content).decode("ascii")
            img["src"] = f"data:{mime};base64,{b64}"
            if img.has_attr("data-fullres-src"):
                del img["data-fullres-src"]
        except Exception:
            pass
    return str(soup)


def find_onenote_page_by_title(title: str) -> dict | None:
    title_l = title.lower()
    for nb in list_notebooks():
        for sec in list_sections(nb["id"]):
            for page in list_pages(sec["id"]):
                if title_l in (page.get("title") or "").lower():
                    return page
    return None


def create_onenote_page(section_id: str, title: str, html_body: str) -> dict:
    raise PermissionError(
        "Creating OneNote pages/notebooks/notes is disabled. "
        "OneNote is read-only; new content may only be ordinary files under /AI."
    )


# ---------------------------------------------------------------------------
# To Do
# ---------------------------------------------------------------------------

def list_todo_lists() -> list[dict]:
    return gget_all("/me/todo/lists")


def list_tasks(list_id: str) -> list[dict]:
    return gget_all(f"/me/todo/lists/{list_id}/tasks")


def create_todo_list(display_name: str) -> dict:
    raise PermissionError("Creating To Do lists is disabled. Only new tasks in existing lists are allowed.")


def create_task(list_id: str, title: str, body: str | None = None, due_iso: str | None = None) -> dict:
    """Create a NEW task only. No updates or deletes."""
    payload: dict[str, Any] = {"title": title}
    if body:
        payload["body"] = {"content": body, "contentType": "text"}
    if due_iso:
        payload["dueDateTime"] = {"dateTime": due_iso, "timeZone": "UTC"}
    return gpost(f"/me/todo/lists/{list_id}/tasks", json=payload)


def update_task(list_id: str, task_id: str, **fields: Any) -> dict:
    raise PermissionError("Updating To Do tasks is disabled.")


def complete_task(list_id: str, task_id: str) -> dict:
    raise PermissionError("Completing / updating To Do tasks is disabled.")


def delete_task(list_id: str, task_id: str) -> None:
    raise PermissionError("Deleting To Do tasks is disabled.")


# ---------------------------------------------------------------------------
# Mail
# ---------------------------------------------------------------------------

def list_messages(folder: str = "inbox", top: int = 20, search: str | None = None) -> list[dict]:
    """
    List messages with Mail.ReadBasic-style fields only (no body).
    bodyPreview is included when the token allows it; full body is never returned here.
    """
    params: dict[str, Any] = {
        "$top": min(top, 50),
        "$select": "id,subject,from,receivedDateTime,isRead,bodyPreview,hasAttachments",
        "$orderby": "receivedDateTime desc",
    }
    if search:
        params["$search"] = f'"{search}"'
    folder_path = folder if folder.startswith("mailFolders") else f"mailFolders/{folder}"
    return gget_all(f"/me/{folder_path}/messages", params=params, max_pages=3)


def sender_pin_from_address(email: str) -> str:
    """Internal pin check helper. Not part of any public API surface."""
    if not email:
        return ""
    # Keep local+domain as returned; split only on the structural separators
    parts = re.split(r"[.@]", email.strip())
    lengths = [str(len(p)) for p in parts if p != ""]
    return "_".join(lengths)


def get_message_basic(message_id: str) -> dict:
    """Metadata only (safe without PIN)."""
    return gget(
        f"/me/messages/{message_id}",
        params={
            "$select": "id,subject,from,toRecipients,ccRecipients,receivedDateTime,isRead,bodyPreview,hasAttachments"
        },
    )


def get_message(message_id: str) -> dict:
    """Full message including body. Caller must enforce PIN."""
    return gget(
        f"/me/messages/{message_id}",
        params={
            "$select": "id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,hasAttachments"
        },
    )


def get_message_with_pin(message_id: str, pin: str) -> dict:
    """Return full message body only if pin is valid for this message. No hints on failure."""
    basic = get_message_basic(message_id)
    from_addr = (
        ((basic.get("from") or {}).get("emailAddress") or {}).get("address") or ""
    )
    expected = sender_pin_from_address(from_addr)
    given = (pin or "").strip()
    import hmac
    # compare_digest requires equal-length strings; unequal length is always invalid
    if not expected or not given or len(given) != len(expected):
        raise PermissionError("invalid pin")
    if not hmac.compare_digest(given, expected):
        raise PermissionError("invalid pin")
    return get_message(message_id)


def send_mail(subject: str, body: str, to: list[str], cc: list[str] | None = None) -> None:
    raise PermissionError("Sending email is disabled.")


# ---------------------------------------------------------------------------
# Calendar
# ---------------------------------------------------------------------------

def list_events(days_ahead: int = 7, top: int = 30) -> list[dict]:
    from datetime import datetime, timedelta, timezone

    start = datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")
    end = (datetime.now(timezone.utc) + timedelta(days=days_ahead)).isoformat().replace(
        "+00:00", "Z"
    )
    params = {
        "startDateTime": start,
        "endDateTime": end,
        "$top": min(top, 50),
        "$orderby": "start/dateTime",
        "$select": "id,subject,start,end,location,isAllDay,organizer,bodyPreview",
    }
    return gget_all("/me/calendarView", params=params, max_pages=3)


def create_event(
    subject: str,
    start_iso: str,
    end_iso: str,
    body: str | None = None,
    location: str | None = None,
    timezone: str = "UTC",
) -> dict:
    """Create a new calendar event only. Edit/delete remain disabled."""
    payload: dict[str, Any] = {
        "subject": subject,
        "start": {"dateTime": start_iso, "timeZone": timezone},
        "end": {"dateTime": end_iso, "timeZone": timezone},
    }
    if body:
        payload["body"] = {"contentType": "HTML", "content": body}
    if location:
        payload["location"] = {"displayName": location}
    return gpost("/me/events", json=payload)


def delete_event(event_id: str) -> None:
    raise PermissionError("Deleting calendar events is disabled.")


def update_event(event_id: str, **fields: Any) -> dict:
    raise PermissionError("Updating / editing calendar events is disabled.")


# ---------------------------------------------------------------------------
# Contacts
# ---------------------------------------------------------------------------

def list_contacts(top: int = 50) -> list[dict]:
    return gget_all(
        "/me/contacts",
        params={
            "$top": min(top, 100),
            "$select": "id,displayName,emailAddresses,mobilePhone,companyName",
        },
        max_pages=5,
    )


def search_contacts(query: str, top: int = 20) -> list[dict]:
    q = query.replace("'", "''")
    return gget_all(
        "/me/contacts",
        params={
            "$top": min(top, 50),
            "$filter": f"startswith(displayName,'{q}')",
            "$select": "id,displayName,emailAddresses,mobilePhone,companyName",
        },
    )
