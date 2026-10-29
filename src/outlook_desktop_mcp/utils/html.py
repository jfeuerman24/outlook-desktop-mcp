"""HTML styling so drafts written by the server match the user's Outlook default font.

Outlook's reply template and plain ``Body`` assignments fall back to the recipient's or
Word's defaults (often Calibri 11). ``draft_email`` and ``draft_reply`` wrap what they write
in an explicit font so it renders the same everywhere. Defaults: Aptos, 12 pt, black.
Override with the settings file (``set_setting``) or environment variables:

    OUTLOOK_MCP_FONT_FAMILY   (default "Aptos")
    OUTLOOK_MCP_FONT_SIZE_PT  (default "12")
    OUTLOOK_MCP_FONT_COLOR    (default "#000000")
    OUTLOOK_MCP_PLAIN_TEXT=1  write plain text instead of styled HTML

Whether a saved draft opens in an Outlook window is the ``open_in_outlook`` setting
(``settings.should_open``), not something this module decides.

Paragraphs are wrapped as ``<p style=…><span style=…>`` because Outlook's Word renderer
ignores inherited font on a bare container; this mirrors what Outlook itself writes.
``text_to_html`` also renders numbered and bulleted lists, ``> quoted`` lines and ``**bold**``;
see its docstring.
"""
import html
import re

from outlook_desktop_mcp.utils import settings

DEFAULT_FAMILY = "Aptos"
DEFAULT_SIZE_PT = "12"
DEFAULT_COLOR = "#000000"

_BODY_OPEN = re.compile(r"<body\b[^>]*>", re.I)


def plain_text_mode() -> bool:
    return bool(settings.get("plain_text"))


def font_style(color: str = "") -> str:
    """The inline CSS for text the server writes; ``color`` overrides the font_color setting."""
    family = str(settings.get("font_family") or DEFAULT_FAMILY).strip() or DEFAULT_FAMILY
    size = str(settings.get("font_size_pt") or DEFAULT_SIZE_PT).strip() or DEFAULT_SIZE_PT
    color = (color or "").strip() or str(settings.get("font_color") or DEFAULT_COLOR).strip() or DEFAULT_COLOR
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


def text_to_html(text: str, color: str = "") -> str:
    """Turn plain text into styled, single-spaced HTML, understanding a little structure.

    Every line is exactly one line (``line-height:100%``, no paragraph spacing) and a
    blank line in the text becomes exactly one blank line in the email, matching what a
    person typing in Outlook gets. Within a paragraph, single newlines become ``<br>``.
    A paragraph whose lines start with ``1.`` / ``2)`` becomes a numbered list, one with
    ``-``, ``*`` or ``\u2022`` a bulleted list; a line without a marker continues the
    previous item on its own line. A paragraph whose lines start with ``>`` is indented
    (useful for text the reader should copy). ``**bold**`` works anywhere. Everything
    else is escaped literally. ``color`` overrides the font_color setting.
    """
    style = font_style(color)
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


# ---------------------------------------------------------------------------
# The note the server writes into a draft is wrapped in a marked <div> so that
# update_draft can find and replace it later without touching the signature or
# the quoted history underneath.
# ---------------------------------------------------------------------------
NOTE_ID = "outlook-mcp-note"
_NOTE_START = re.compile(r"<div\b[^>]*\bid\s*=\s*[\"']?" + NOTE_ID + r"[\"']?[^>]*>", re.I)
_DIV_TOKEN = re.compile(r"<(/?)div\b[^>]*>", re.I)


def wrap_note(fragment: str) -> str:
    return f'<div id="{NOTE_ID}">{fragment}</div>'


def find_note(document: str) -> tuple[int, int] | None:
    """``(start, end)`` of the marked note div in ``document`` (end is past the closing tag),
    found by balancing nested <div> tags; None when there is no marker."""
    m = _NOTE_START.search(document or "")
    if not m:
        return None
    depth = 1
    for tok in _DIV_TOKEN.finditer(document, m.end()):
        depth += -1 if tok.group(1) else 1
        if depth == 0:
            return m.start(), tok.end()
    return None


def replace_note(document: str, fragment: str) -> str | None:
    """``document`` with the marked note replaced by ``wrap_note(fragment)``; None if no marker."""
    span = find_note(document)
    if span is None:
        return None
    start, end = span
    return document[:start] + wrap_note(fragment) + document[end:]


# ---------------------------------------------------------------------------
# Inline answers: text written inside the quoted message, under the paragraph it
# answers, in its own colour, the way people answer a numbered list point by point.
# An anchor is a phrase as the reader sees it; it is looked for in the visible text
# of the HTML, so tags, entities and Word's curly quotes in between do not matter.
# Every block written carries INLINE_CLASS so update_draft can remove and rewrite them.
# ---------------------------------------------------------------------------
INLINE_CLASS = "outlook-mcp-inline"
DEFAULT_INLINE_COLOR = "#C00000"

_TOKEN = re.compile(
    r"<!--.*?-->|<(style|script|head)\b.*?</\1\s*>|<[^>]*>|&[#A-Za-z0-9]+;|[^<&]+|&",
    re.I | re.S,
)
_BREAK_TAG = re.compile(r"</?(p|div|li|tr|td|th|h[1-6]|ul|ol|table|blockquote|pre)\b|<br\b", re.I)
_BLOCK_END = re.compile(r"</(p|div|h[1-6])\s*>|</(li|td)\s*>|<br\b[^>]*>", re.I)
_BODY_CLOSE = re.compile(r"</body\s*>", re.I)
_PUNCT = str.maketrans({"‘": "'", "’": "'", "“": '"', "”": '"',
                        "–": "-", "—": "-", "\xa0": " "})
_INLINE_BLOCK = re.compile(
    r"\s*<(p|ol|ul|span)\b[^>]*\bclass\s*=\s*[\"']?" + INLINE_CLASS + r"\b[^>]*>.*?</\1\s*>(?:\s*<br\b[^>]*>)?\s*",
    re.I | re.S,
)


def normalize_anchor(text: str) -> str:
    """Lower-case, straight quotes, single spaces: how anchors and visible text are compared."""
    return re.sub(r"\s+", " ", (text or "").translate(_PUNCT)).strip().lower()


def _visible_text(document: str, start: int) -> tuple[str, list[int]]:
    """The text a reader sees in ``document[start:]``, normalised like ``normalize_anchor``,
    and for each character the offset in ``document`` just past its source."""
    chars: list[str] = []
    ends: list[int] = []

    def emit(ch: str, end: int) -> None:
        ch = ch.translate(_PUNCT)
        if ch.isspace():
            if not chars or chars[-1] == " ":
                return
            ch = " "
        chars.append(ch.lower())
        ends.append(end)

    for m in _TOKEN.finditer(document, start):
        tok = m.group(0)
        if tok.startswith("<"):
            if _BREAK_TAG.match(tok):
                emit(" ", m.end())
            continue
        if tok.startswith("&") and len(tok) > 1:
            for ch in html.unescape(tok):
                emit(ch, m.end())
            continue
        for i, ch in enumerate(tok):
            emit(ch, m.start() + i + 1)
    return "".join(chars), ends


def find_anchor(document: str, anchor: str, start: int = 0) -> int | None:
    """Offset in ``document`` just past the first occurrence of ``anchor`` at or after
    ``start``, matched on the visible text; None when it is not there."""
    needle = normalize_anchor(anchor)
    if not needle:
        return None
    text, ends = _visible_text(document, start)
    idx = text.find(needle)
    if idx < 0:
        return None
    return ends[idx + len(needle) - 1]


def inline_html(text: str, color: str) -> tuple[str, str]:
    """Two renderings of one inline answer, both tagged with INLINE_CLASS: a block (styled
    paragraphs, for after a paragraph) and a line (a span and a <br>, for after a <br>)."""
    style = font_style(color)
    block = re.sub(r"<(p|ol|ul) style=", rf'<\1 class="{INLINE_CLASS}" style=', text_to_html(text, color))
    lines = [_inline(ln.strip()) for ln in (text or "").replace("\r\n", "\n").replace("\r", "\n").split("\n")]
    line = f'<span class="{INLINE_CLASS}" style="{style}">' + "<br>".join(lines) + "</span><br>"
    return block, line


def insert_inline(document: str, answers, color: str = DEFAULT_INLINE_COLOR) -> tuple[str, list[str]]:
    """Write each ``(anchor, text)`` answer into ``document`` after the block that contains
    the anchor: after the closing ``</p>``/``</div>``, inside a ``</li>``/``</td>``, or as a
    line after a ``<br>``. Each search starts where the previous answer went in, so a
    phrase that repeats resolves in document order. Returns the new document and the
    anchors not found; when any is missing the document comes back unchanged."""
    out = document or ""
    cursor = 0
    missing: list[str] = []
    for anchor, text in answers:
        pos = find_anchor(out, anchor, cursor)
        if pos is None:
            missing.append(anchor)
            continue
        block, line = inline_html(text, color)
        m = _BLOCK_END.search(out, pos)
        if m is None:
            end = _BODY_CLOSE.search(out, pos)
            insert_at, fragment = (end.start() if end else len(out)), block
        elif m.group(1):
            insert_at, fragment = m.end(), block
        elif m.group(2):
            insert_at, fragment = m.start(), block
        else:
            insert_at, fragment = m.end(), line
        out = out[:insert_at] + fragment + out[insert_at:]
        cursor = insert_at + len(fragment)
    if missing:
        return document or "", missing
    return out, []


def strip_inline(document: str) -> str:
    """``document`` without the inline answers written earlier (by INLINE_CLASS)."""
    return _INLINE_BLOCK.sub("", document or "")
