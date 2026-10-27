"""Plain-text body cleanup: drop the quoted history under a reply and Outlook's external-sender banner.

Outlook's ``Body`` is plain text, so a reply carries the whole earlier chain under it.
When a client reads several messages of one thread that history is repeated per message
and the 5 000-character cap then cuts the *newest* text. ``trim_quoted_history`` cuts at
the first quote marker instead, so the cap applies to the message's own words.

Pure functions, unit-tested without Outlook.
"""
import re

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


def clean_body(text: str, *, trim_history: bool = True, max_length: int = 5000) -> dict:
    """``{"body", "body_trimmed", "body_truncated"}`` for a plain-text Outlook body."""
    if trim_history:
        body, trimmed = trim_quoted_history(text)
    else:
        body, trimmed = tidy(strip_external_banner(text)), False
    truncated = len(body) > max_length
    if truncated:
        body = body[:max_length].rstrip() + "\n... [truncated]"
    return {"body": body, "body_trimmed": trimmed, "body_truncated": truncated}
