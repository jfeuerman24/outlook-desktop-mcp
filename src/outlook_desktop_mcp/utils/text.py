"""Plain-text body cleanup: drop the quoted history under a reply, Outlook's external-sender
banner, and the tracking links that make up most of a long email.

Outlook's ``Body`` is plain text, so a reply carries the whole earlier chain under it.
When a client reads several messages of one thread that history is repeated per message
and the character cap then cuts the *newest* text. ``trim_quoted_history`` cuts at the
first quote marker instead, so the cap applies to the message's own words.

Links are the other thing that fills a body. Outlook renders every HTML link as
``text <https://...>``, and in confirmations and newsletters the URLs are 85 to 99 per cent
of the characters (a 57,000-character flight confirmation held 6,500 of text). ``shorten_links``
turns each long URL into ``[link: domain]``, drops links to image files, and collapses
repeats, so the cap falls on words rather than tracking parameters.

Pure functions, unit-tested without Outlook.
"""
import re
from urllib.parse import urlparse

DEFAULT_BODY_CHARS = 8000  # cap on a body after trimming and link shortening
KEEP_URL_CHARS = 60  # a URL this short is left as it is; it may be the point of the message
_IMAGE_EXTENSIONS = (".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg", ".bmp")
_URL = re.compile(r"<(https?://[^>\s]+)>|(https?://[^\s<>\"']+)")
_LINK_RUNS = re.compile(r"(\[link: [^\]]*\])(?:[ \t]*\n?[ \t]*\1)+")
_SPACES = re.compile(r"[ \t]{3,}")

BLANK_CHARS = " \t\r\n\xa0​﻿"

_BANNER = re.compile(
    r"You don.t often get email from.*?(?:Learn why this is important\S*|\n)", re.I | re.S
)

# Where the quoted original begins. Each marker must sit at a line start.
_QUOTE_MARKERS = (
    # Outlook: "From: X" header block, with a Sent/Date/To line within the next few lines
    re.compile(r"^[ \t]*From:[^\n]*\n(?:[^\n]*\n){0,3}?[ \t]*(?:Sent|Date|To):", re.M),
    re.compile(r"^[ \t]*-{2,}[ \t]*Original (?:Message|Appointment)[ \t]*-{2,}", re.I | re.M),
    # Gmail / Apple Mail: "On Mon, Oct 26, 2026 at 5:40 PM Linda <linda@…> wrote:" (may wrap)
    re.compile(r"^[ \t]*On .{0,300}?wrote:[ \t]*$", re.I | re.M | re.S),
    # Outlook's separator line above a quoted message
    re.compile(r"^[ \t]*_{5,}[ \t]*$", re.M),
)


def normalize_newlines(text: str) -> str:
    return (text or "").replace("\r\n", "\n").replace("\r", "\n")


def strip_external_banner(text: str) -> str:
    return _BANNER.sub("", normalize_newlines(text))


def find_quote_start(text: str) -> int | None:
    """Offset of the earliest quote marker in ``text`` (already newline-normalized), or None."""
    cut: int | None = None
    for pattern in _QUOTE_MARKERS:
        match = pattern.search(text)
        if match and (cut is None or match.start() < cut):
            cut = match.start()
    return cut


def tidy(text: str) -> str:
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def trim_quoted_history(text: str) -> tuple[str, bool]:
    """Return ``(own text, was_trimmed)``.

    Bare-forward exception: when nothing meaningful precedes the marker, the quoted part
    *is* the content, so the whole body is kept.
    """
    text = strip_external_banner(text)
    cut = find_quote_start(text)
    if cut is None or not text[:cut].strip(BLANK_CHARS):
        return tidy(text), False
    return tidy(text[:cut]), True


def _domain(url: str) -> str:
    try:
        host = urlparse(url).netloc.lower().split(":")[0]
    except Exception:
        return ""
    parts = host.split(".")
    return ".".join(parts[-2:]) if len(parts) >= 2 else host


def shorten_links(text: str) -> tuple[str, int]:
    """``(text, count)`` with every URL longer than ``KEEP_URL_CHARS`` replaced by
    ``[link: domain]`` (dropped outright when it points at an image file), runs of the
    same placeholder collapsed to one, and runs of spaces squeezed. Short URLs stay."""
    count = 0

    def replace(match):
        nonlocal count
        url = match.group(1) or match.group(2)
        if len(url) <= KEEP_URL_CHARS:
            return match.group(0)
        count += 1
        path = url.split("?", 1)[0].split("#", 1)[0].lower()
        if path.endswith(_IMAGE_EXTENSIONS):
            return ""
        return f"[link: {_domain(url)}]"

    out = _URL.sub(replace, text or "")
    if count:
        out = _LINK_RUNS.sub(r"\1", out)
        out = _SPACES.sub("  ", out)
    return out, count


def clean_body(text: str, *, trim_history: bool = True, max_length: int = DEFAULT_BODY_CHARS) -> dict:
    """``{"body", "body_trimmed", "body_truncated"}`` for a plain-text Outlook body, plus
    ``links_shortened`` when links were, and ``body_total_chars`` when the cap cut it
    (so the caller knows how much to ask for). ``max_length`` 0 means no cap."""
    if trim_history:
        body, trimmed = trim_quoted_history(text)
    else:
        body, trimmed = tidy(strip_external_banner(text)), False
    body, links = shorten_links(body)
    if links:
        body = tidy(body)
    total = len(body)
    truncated = bool(max_length) and total > max_length
    if truncated:
        body = body[:max_length].rstrip() + "\n... [truncated]"
    out = {"body": body, "body_trimmed": trimmed, "body_truncated": truncated}
    if links:
        out["links_shortened"] = links
    if truncated:
        out["body_total_chars"] = total
    return out
