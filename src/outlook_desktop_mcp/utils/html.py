"""Outgoing HTML styling so mail sent through the server matches the user's Outlook default font.

Outlook's reply template and plain ``Body`` assignments fall back to the recipient's or
Word's defaults (often Calibri 11). ``send_email`` and ``reply_email`` wrap what they send
in an explicit font so it renders the same everywhere. Defaults: Aptos, 12 pt, black.
Override with environment variables:

    OUTLOOK_MCP_FONT_FAMILY   (default "Aptos")
    OUTLOOK_MCP_FONT_SIZE_PT  (default "12")
    OUTLOOK_MCP_FONT_COLOR    (default "#000000")
    OUTLOOK_MCP_PLAIN_TEXT=1  send plain text instead of styled HTML

Paragraphs are wrapped as ``<p style=…><span style=…>`` because Outlook's Word renderer
ignores inherited font on a bare container; this mirrors what Outlook itself writes.
"""
import html
import os
import re

DEFAULT_FAMILY = "Aptos"
DEFAULT_SIZE_PT = "12"
DEFAULT_COLOR = "#000000"

_BODY_OPEN = re.compile(r"<body\b[^>]*>", re.I)


def plain_text_mode() -> bool:
    return os.environ.get("OUTLOOK_MCP_PLAIN_TEXT", "").strip().lower() in {"1", "true", "yes", "on"}


def font_style() -> str:
    family = os.environ.get("OUTLOOK_MCP_FONT_FAMILY", DEFAULT_FAMILY).strip() or DEFAULT_FAMILY
    size = os.environ.get("OUTLOOK_MCP_FONT_SIZE_PT", DEFAULT_SIZE_PT).strip() or DEFAULT_SIZE_PT
    color = os.environ.get("OUTLOOK_MCP_FONT_COLOR", DEFAULT_COLOR).strip() or DEFAULT_COLOR
    return f"font-family:{family},Calibri,sans-serif;font-size:{size}pt;color:{color}"


def text_to_html(text: str) -> str:
    """Escape plain text; blank lines become paragraphs, single newlines become ``<br>``."""
    style = font_style()
    normalized = (text or "").replace("\r\n", "\n").replace("\r", "\n").strip()
    paragraphs = re.split(r"\n\s*\n", normalized) if normalized else [""]
    out = []
    for paragraph in paragraphs:
        lines = html.escape(paragraph.strip()).replace("\n", "<br>")
        out.append(f'<p style="margin:0 0 12pt 0;{style}"><span style="{style}">{lines}</span></p>')
    return "\n".join(out)


def style_html(fragment: str) -> str:
    """Wrap caller-supplied HTML so unstyled text inherits the default font."""
    return f'<div style="{font_style()}">{fragment}</div>'


def inject_into_body(document: str, fragment: str) -> str:
    """Insert ``fragment`` right after ``<body>`` so it sits above Outlook's quoted history."""
    match = _BODY_OPEN.search(document or "")
    if not match:
        return fragment + (document or "")
    end = match.end()
    return document[:end] + "\n" + fragment + "\n" + document[end:]
