"""HTML styling so drafts written by the server match the user's Outlook default font.

Outlook's reply template and plain ``Body`` assignments fall back to the recipient's or
Word's defaults (often Calibri 11). ``draft_email`` and ``draft_reply`` wrap what they write
in an explicit font so it renders the same everywhere. Defaults: Aptos, 12 pt, black.
Override with environment variables:

    OUTLOOK_MCP_FONT_FAMILY   (default "Aptos")
    OUTLOOK_MCP_FONT_SIZE_PT  (default "12")
    OUTLOOK_MCP_FONT_COLOR    (default "#000000")
    OUTLOOK_MCP_PLAIN_TEXT=1  write plain text instead of styled HTML
    OUTLOOK_MCP_OPEN_DRAFTS=1 open each new draft in an Outlook window (default: just
                              save it to Drafts; the per-call open_in_outlook flag
                              still overrides either way)

Paragraphs are wrapped as ``<p style=…><span style=…>`` because Outlook's Word renderer
ignores inherited font on a bare container; this mirrors what Outlook itself writes.
``text_to_html`` also renders numbered and bulleted lists, ``> quoted`` lines and ``**bold**``;
see its docstring.
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


def open_drafts_default() -> bool:
    """Whether draft_* tools open the saved item in an Outlook window when not told."""
    return os.environ.get("OUTLOOK_MCP_OPEN_DRAFTS", "").strip().lower() in {"1", "true", "yes", "on"}


def font_style() -> str:
    family = os.environ.get("OUTLOOK_MCP_FONT_FAMILY", DEFAULT_FAMILY).strip() or DEFAULT_FAMILY
    size = os.environ.get("OUTLOOK_MCP_FONT_SIZE_PT", DEFAULT_SIZE_PT).strip() or DEFAULT_SIZE_PT
    color = os.environ.get("OUTLOOK_MCP_FONT_COLOR", DEFAULT_COLOR).strip() or DEFAULT_COLOR
    return f"font-family:{family},Calibri,sans-serif;font-size:{size}pt;color:{color}"


_OL_ITEM = re.compile(r"^\s*(\d+)[.)]\s+(.*)$")
_UL_ITEM = re.compile(r"^\s*[-*\u2022]\s+(.*)$")
_QUOTE = re.compile(r"^\s*>\s?(.*)$")
_BOLD = re.compile(r"\*\*(.+?)\*\*")


def _inline(text: str) -> str:
    """Escape one line and turn **bold** into <b>."""
    escaped = html.escape(text)
    return _BOLD.sub(r"<b>\1</b>", escaped)


def _item_html(lines: list[str]) -> str:
    """Join the lines of a list item; continuation lines sit under the item."""
    return "<br>".join(_inline(line.strip()) for line in lines if line.strip())


def _blank(style: str) -> str:
    """One empty single-spaced line, used between blocks."""
    return f'<p style="margin:0;line-height:100%;{style}"><span style="{style}">&nbsp;</span></p>'


def text_to_html(text: str) -> str:
    """Turn plain text into styled, single-spaced HTML, understanding a little structure.

    Every line is exactly one line (``line-height:100%``, no paragraph spacing) and a
    blank line in the text becomes exactly one blank line in the email, matching what a
    person typing in Outlook gets. Within a paragraph, single newlines become ``<br>``.
    A paragraph whose lines start with ``1.`` / ``2)`` becomes a numbered list, one with
    ``-``, ``*`` or ``\u2022`` a bulleted list; a line without a marker continues the
    previous item on its own line. A paragraph whose lines start with ``>`` is indented
    (useful for text the reader should copy). ``**bold**`` works anywhere. Everything
    else is escaped literally.
    """
    style = font_style()
    normalized = (text or "").replace("\r\n", "\n").replace("\r", "\n").strip()
    paragraphs = re.split(r"\n\s*\n", normalized) if normalized else [""]
    line = f"margin:0;line-height:100%;{style}"
    span = f'<span style="{style}">'
    blocks = []
    for paragraph in paragraphs:
        lines = paragraph.strip().split("\n")
        first = lines[0]
        if _OL_ITEM.match(first) or _UL_ITEM.match(first):
            ordered = bool(_OL_ITEM.match(first))
            pattern = _OL_ITEM if ordered else _UL_ITEM
            items: list[list[str]] = []
            for ln in lines:
                m = pattern.match(ln)
                if m:
                    items.append([m.group(2) if ordered else m.group(1)])
                elif items:
                    items[-1].append(ln)
            tag = "ol" if ordered else "ul"
            li = "".join(f'<li style="{line}">{span}{_item_html(item)}</span></li>' for item in items)
            blocks.append(f'<{tag} style="margin:0;padding-left:24pt;{line}">{li}</{tag}>')
        elif all(_QUOTE.match(ln) for ln in lines):
            body = "<br>".join(_inline(_QUOTE.match(ln).group(1)) for ln in lines)
            blocks.append(f'<p style="{line};margin-left:24pt">{span}{body}</span></p>')
        else:
            body = "<br>".join(_inline(ln.strip()) for ln in lines)
            blocks.append(f'<p style="{line}">{span}{body}</span></p>')
    return ("\n" + _blank(style) + "\n").join(blocks)


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
