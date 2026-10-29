"""Filter strings for Outlook's ``Items.Restrict``. Pure Python, unit-tested.

Two dialects exist. Jet (``[ReceivedTime] >= '...'``) is simple but cannot do substring
matches. DASL (``@SQL="urn:..." LIKE '%x%'``) can, and Outlook evaluates it itself, so a
sender or recipient filter that would take a minute walking items from Python takes a
fraction of a second here. The two can be chained: ``items.Restrict(jet).Restrict(dasl)``.

Property names are the ones Outlook accepts in Restrict. ``PR_SENDER_SMTP_ADDRESS`` is the
message's own copy of the sender's SMTP address; inside an Exchange organisation
``fromemail`` holds a directory name instead, so both are searched.
"""
from __future__ import annotations

from datetime import datetime

FROM_NAME = "urn:schemas:httpmail:fromname"
FROM_EMAIL = "urn:schemas:httpmail:fromemail"
SENDER_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"
SUBJECT = "urn:schemas:httpmail:subject"
BODY = "urn:schemas:httpmail:textdescription"
DISPLAY_TO = "urn:schemas:httpmail:displayto"
DISPLAY_CC = "urn:schemas:httpmail:displaycc"
DISPLAY_BCC = "urn:schemas:httpmail:displaybcc"
DATE_RECEIVED = "urn:schemas:httpmail:datereceived"
READ = "urn:schemas:httpmail:read"
HAS_ATTACHMENT = "urn:schemas:httpmail:hasattachment"
# A follow-up flag is a task on the message (Outlook 2007+): PidLidTaskComplete is false while
# the flag is open and true once it is marked complete. PR_FLAG_STATUS = 2 is the older,
# pre-task form of the same flag and is kept for messages written that way.
TASK_COMPLETE = "http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/811C000B"
FLAG_STATUS = "http://schemas.microsoft.com/mapi/proptag/0x10900003"

SENDER_PROPS = (FROM_NAME, FROM_EMAIL, SENDER_SMTP)
RECIPIENT_PROPS = (DISPLAY_TO, DISPLAY_CC, DISPLAY_BCC)


def escape_like(text: str) -> str:
    """Make ``text`` safe inside a LIKE pattern: wildcards and brackets become literals,
    quotes are doubled as DASL requires."""
    text = text.replace("[", "[[]").replace("%", "[%]").replace("_", "[_]")
    return text.replace("'", "''").replace('"', '""')


def escape_quotes(text: str) -> str:
    return text.replace("'", "''").replace('"', '""')


def jet_date(when: datetime) -> str:
    """The date format Outlook's filters accept."""
    return when.strftime("%m/%d/%Y %H:%M")


def contains(prop: str, needle: str) -> str:
    return f"\"{prop}\" LIKE '%{escape_like(needle)}%'"


def any_contains(props, needles) -> str:
    """OR of every (property, needle) pair, parenthesised."""
    if isinstance(needles, str):
        needles = [needles]
    parts = [contains(p, n) for n in needles for p in props if n]
    if not parts:
        raise ValueError("no needle")
    return "(" + " OR ".join(parts) + ")"


def sql(*clauses: str) -> str:
    """Join DASL clauses with AND and add the ``@SQL=`` prefix Restrict needs."""
    parts = [c for c in clauses if c]
    if not parts:
        return ""
    return "@SQL=" + " AND ".join(parts)


def sender_filter(needle: str) -> str:
    """Messages whose sender name, directory address or SMTP address contains ``needle``."""
    return any_contains(SENDER_PROPS, needle)


def recipient_filter(needles) -> str:
    """Messages whose To, CC or BCC display string contains any of ``needles``.

    Display strings hold names ("Lou Castelli"), and an address only when Outlook had no
    name for it, so pass the names the person is known by as well as the needle itself.
    """
    if isinstance(needles, str):
        needles = [needles]
    seen: list[str] = []
    for n in needles:
        n = (n or "").strip()
        if len(n) >= 2 and n.lower() not in (s.lower() for s in seen):
            seen.append(n)
    return any_contains(RECIPIENT_PROPS, seen)


def subject_filter(needle: str) -> str:
    return contains(SUBJECT, needle)


def people_filter(needle: str, names=None) -> str:
    """Messages from or to ``needle``: the sender's name or address, or the To/CC/BCC
    display strings (plus any ``names`` the person is known by). A sender filter alone
    misses mail where the person is a co-recipient, such as a colleague writing to them
    with the user copied, and the user's own replies filed outside Sent Items."""
    return "(" + sender_filter(needle) + " OR " + recipient_filter([needle] + list(names or [])) + ")"


def anywhere_filter(needle: str, indexed_body: bool | None = None) -> str:
    """Subject, sender or recipients containing ``needle``, in one filter Outlook
    evaluates itself. With ``indexed_body`` not None the body is included too
    (True = search-index phrase match, False = substring scan of every body)."""
    parts = [subject_filter(needle), sender_filter(needle)]
    if len((needle or "").strip()) >= 2:
        parts.append(recipient_filter([needle]))
    if indexed_body is not None:
        parts.append(body_filter(needle, indexed_body))
    return "(" + " OR ".join(parts) + ")"


def body_filter(needle: str, indexed: bool = False) -> str:
    """Body text containing ``needle``. ``indexed`` uses the search index's phrase match
    (whole words, fast) instead of a substring scan of every body (exact, slow)."""
    if indexed:
        return f"\"{BODY}\" ci_phrasematch '{escape_quotes(needle)}'"
    return contains(BODY, needle)


def subject_or_body_filter(needle: str, indexed_body: bool = False) -> str:
    return "(" + subject_filter(needle) + " OR " + body_filter(needle, indexed_body) + ")"


def received_since(when: datetime) -> str:
    return f"\"{DATE_RECEIVED}\" >= '{jet_date(when)}'"


def received_until(when: datetime) -> str:
    return f"\"{DATE_RECEIVED}\" <= '{jet_date(when)}'"


def unread_filter() -> str:
    return f"\"{READ}\" = 0"


def has_attachment_filter() -> str:
    """Messages with at least one attachment (inline pictures count, so a name match
    in Python still follows)."""
    return f"\"{HAS_ATTACHMENT}\" = 1"


def flagged_filter() -> str:
    """Messages flagged for follow-up and not yet marked complete."""
    return f"(\"{TASK_COMPLETE}\" = 0 OR \"{FLAG_STATUS}\" = 2)"


def looks_like_address(needle: str) -> bool:
    """True for "lou@pennspring.com", "pennspring.com", "sanjohninc.com"; false for names."""
    n = (needle or "").strip()
    return "@" in n or ("." in n and " " not in n)
