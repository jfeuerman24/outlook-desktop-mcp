"""Pull readable text out of a file so Claude can read an attachment without the user
opening it. Pure Python: pypdf, python-docx, openpyxl, python-pptx. Each extractor
returns ``(text, info)`` where ``info`` describes what was read (pages, sheets, ...).

Text is capped at ``max_chars``; ``info["truncated"]`` says when the cap was hit, so the
caller can ask for a narrower range (``pages`` / ``sheet``) instead of the whole file.
Scanned PDFs (pictures of pages) contain no text layer and come back empty; that is
reported rather than guessed at. Old binary Office formats (.doc, .xls, .ppt) are not
supported.
"""
from __future__ import annotations

import csv
import io
import pathlib
import re

SUPPORTED = {
    ".pdf": "pdf", ".docx": "docx", ".xlsx": "xlsx", ".xlsm": "xlsx", ".pptx": "pptx",
    ".csv": "csv", ".tsv": "csv", ".txt": "text", ".md": "text", ".json": "text",
    ".log": "text", ".xml": "text", ".html": "html", ".htm": "html", ".eml": "text",
}
DEFAULT_MAX_CHARS = 20000
MAX_SHEET_ROWS = 500


def _cap(text: str, max_chars: int) -> tuple[str, bool]:
    if max_chars and len(text) > max_chars:
        return text[:max_chars], True
    return text, False


def parse_pages(spec: str, total: int) -> list[int]:
    """'1-3,7' -> [0, 1, 2, 6] (0-based, clipped to ``total``); '' -> every page."""
    if not spec or not spec.strip():
        return list(range(total))
    pages: list[int] = []
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            lo, hi = part.split("-", 1)
            lo_i = int(lo) if lo.strip() else 1
            hi_i = int(hi) if hi.strip() else total
        else:
            lo_i = hi_i = int(part)
        for n in range(max(1, lo_i), min(total, hi_i) + 1):
            if n - 1 not in pages:
                pages.append(n - 1)
    return pages


def extract_pdf(path: pathlib.Path, max_chars: int, pages: str = "") -> tuple[str, dict]:
    from pypdf import PdfReader

    reader = PdfReader(str(path))
    if reader.is_encrypted:
        try:
            reader.decrypt("")
        except Exception:
            return "", {"kind": "pdf", "error": "PDF is password-protected"}
    total = len(reader.pages)
    wanted = parse_pages(pages, total)
    chunks = []
    empty = 0
    for i in wanted:
        try:
            text = reader.pages[i].extract_text() or ""
        except Exception:
            text = ""
        if not text.strip():
            empty += 1
        chunks.append(f"--- page {i + 1} ---\n{text.strip()}")
    text, truncated = _cap("\n\n".join(chunks), max_chars)
    info = {"kind": "pdf", "pages_total": total, "pages_read": [i + 1 for i in wanted], "truncated": truncated}
    if wanted and empty == len(wanted):
        info["note"] = "No text layer on the pages read: this is probably a scanned PDF (pictures of pages). OCR is not available."
    return text, info


def extract_docx(path: pathlib.Path, max_chars: int) -> tuple[str, dict]:
    import docx

    document = docx.Document(str(path))
    parts: list[str] = []
    for para in document.paragraphs:
        if para.text.strip():
            style = (para.style.name or "") if para.style is not None else ""
            prefix = "# " if style.startswith("Heading") or style == "Title" else ""
            parts.append(prefix + para.text.strip())
    tables = 0
    for table in document.tables:
        tables += 1
        rows = []
        for row in table.rows:
            cells = [c.text.strip().replace("\n", " ") for c in row.cells]
            rows.append(" | ".join(cells))
        parts.append(f"--- table {tables} ---\n" + "\n".join(rows))
    text, truncated = _cap("\n\n".join(parts), max_chars)
    return text, {"kind": "docx", "paragraphs": len(document.paragraphs), "tables": tables, "truncated": truncated}


def extract_xlsx(path: pathlib.Path, max_chars: int, sheet: str = "") -> tuple[str, dict]:
    from openpyxl import load_workbook

    # A read-only workbook keeps the file open until close(). Close it on every
    # path (early return, exception too), or Windows cannot delete the temp copy
    # the caller made.
    wb = load_workbook(str(path), read_only=True, data_only=True)
    try:
        names = wb.sheetnames
        if sheet:
            chosen = [n for n in names if n.lower() == sheet.lower()] or [n for n in names if sheet.lower() in n.lower()]
            if not chosen:
                return "", {"kind": "xlsx", "sheets": names, "error": f"No sheet matching {sheet!r}"}
        else:
            chosen = names
        parts = []
        rows_read = {}
        for name in chosen:
            ws = wb[name]
            lines = []
            count = 0
            for row in ws.iter_rows(values_only=True):
                if count >= MAX_SHEET_ROWS:
                    lines.append(f"... (stopped after {MAX_SHEET_ROWS} rows; ask for this sheet alone or a narrower range)")
                    break
                if row is None or all(v is None or str(v).strip() == "" for v in row):
                    continue
                cells = ["" if v is None else str(v) for v in row]
                while cells and cells[-1] == "":
                    cells.pop()
                lines.append("\t".join(cells))
                count += 1
            rows_read[name] = count
            parts.append(f"--- sheet: {name} ---\n" + "\n".join(lines))
    finally:
        wb.close()
    text, truncated = _cap("\n\n".join(parts), max_chars)
    return text, {"kind": "xlsx", "sheets": names, "sheets_read": chosen, "rows_read": rows_read,
                  "truncated": truncated, "note": "Values are cached results; formulas are not shown."}


def extract_pptx(path: pathlib.Path, max_chars: int, pages: str = "") -> tuple[str, dict]:
    from pptx import Presentation

    prs = Presentation(str(path))
    slides = list(prs.slides)
    wanted = parse_pages(pages, len(slides))
    parts = []
    for i in wanted:
        slide = slides[i]
        lines = []
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False) and shape.text_frame is not None:
                for para in shape.text_frame.paragraphs:
                    t = "".join(run.text for run in para.runs).strip()
                    if t:
                        lines.append(t)
            if getattr(shape, "has_table", False) and shape.has_table:
                for row in shape.table.rows:
                    lines.append(" | ".join(c.text.strip().replace("\n", " ") for c in row.cells))
        try:
            notes = slide.notes_slide.notes_text_frame.text.strip() if slide.has_notes_slide else ""
        except Exception:
            notes = ""
        if notes:
            lines.append(f"[notes] {notes}")
        parts.append(f"--- slide {i + 1} ---\n" + "\n".join(lines))
    text, truncated = _cap("\n\n".join(parts), max_chars)
    return text, {"kind": "pptx", "slides_total": len(slides), "slides_read": [i + 1 for i in wanted], "truncated": truncated}


def extract_csv(path: pathlib.Path, max_chars: int) -> tuple[str, dict]:
    raw = path.read_text(encoding="utf-8-sig", errors="replace")
    dialect = csv.excel_tab if path.suffix.lower() == ".tsv" else csv.excel
    lines = []
    count = 0
    for row in csv.reader(io.StringIO(raw), dialect):
        if count >= MAX_SHEET_ROWS:
            lines.append(f"... (stopped after {MAX_SHEET_ROWS} rows)")
            break
        lines.append("\t".join(row))
        count += 1
    text, truncated = _cap("\n".join(lines), max_chars)
    return text, {"kind": "csv", "rows_read": count, "truncated": truncated}


def extract_text_file(path: pathlib.Path, max_chars: int) -> tuple[str, dict]:
    raw = path.read_text(encoding="utf-8-sig", errors="replace")
    text, truncated = _cap(raw, max_chars)
    return text, {"kind": "text", "chars_total": len(raw), "truncated": truncated}


_TAG = re.compile(r"<[^>]+>")


def extract_html_file(path: pathlib.Path, max_chars: int) -> tuple[str, dict]:
    raw = path.read_text(encoding="utf-8-sig", errors="replace")
    raw = re.sub(r"(?is)<(script|style).*?</\1>", " ", raw)
    plain = re.sub(r"\n{3,}", "\n\n", _TAG.sub(" ", raw).replace("&nbsp;", " "))
    plain = re.sub(r"[ \t]+", " ", plain).strip()
    text, truncated = _cap(plain, max_chars)
    return text, {"kind": "html", "truncated": truncated}


def extract(path: str | pathlib.Path, max_chars: int = DEFAULT_MAX_CHARS, pages: str = "",
            sheet: str = "") -> tuple[str, dict]:
    """Dispatch on extension. Unsupported types return ("", {"error": ...})."""
    p = pathlib.Path(path)
    kind = SUPPORTED.get(p.suffix.lower())
    if kind is None:
        return "", {"kind": "unsupported", "error": f"Cannot extract text from {p.suffix or 'a file without extension'}. "
                    f"Supported: {', '.join(sorted(SUPPORTED))}. Use save_attachment to save it instead."}
    try:
        if kind == "pdf":
            return extract_pdf(p, max_chars, pages)
        if kind == "docx":
            return extract_docx(p, max_chars)
        if kind == "xlsx":
            return extract_xlsx(p, max_chars, sheet)
        if kind == "pptx":
            return extract_pptx(p, max_chars, pages)
        if kind == "csv":
            return extract_csv(p, max_chars)
        if kind == "html":
            return extract_html_file(p, max_chars)
        return extract_text_file(p, max_chars)
    except ImportError as e:
        return "", {"kind": kind, "error": f"Missing library for {p.suffix}: {e}. Run install.cmd (or install.ps1) to add it."}
    except Exception as e:
        return "", {"kind": kind, "error": f"Could not read {p.name}: {type(e).__name__}: {e}"}
