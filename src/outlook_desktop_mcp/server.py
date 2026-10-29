"""
Outlook Desktop MCP Server
===========================
Exposes Microsoft Outlook Desktop (Classic) as an MCP server over stdio.
Uses COM automation — no Microsoft Graph, no Entra app registration.
Just run this on Windows with Outlook open and you have an email MCP server.

Nothing here ever sends. ``draft_email``, ``draft_reply`` and ``draft_meeting`` save to
the Drafts folder (or the calendar, unsent) and optionally open the item in Outlook so the
user can check it and press Send themselves. Responding to invitations and cancelling
meetings you organised are not offered, because both would send a message.

Entry point: python -m outlook_desktop_mcp
"""
import sys
import json
import logging
import re

from mcp.server.fastmcp import FastMCP

from outlook_desktop_mcp.com_bridge import OutlookBridge
from datetime import datetime, timedelta

import os
import platform
import shutil
import tempfile
import time

from outlook_desktop_mcp import __version__

from outlook_desktop_mcp.tools import calendars as calendars_mod
from outlook_desktop_mcp.tools.calendars import describe_calendars, is_people_calendar, resolve_calendars
from outlook_desktop_mcp.tools.conversation import conversation_items, folder_path
from outlook_desktop_mcp.utils import dasl, deadline, settings
from outlook_desktop_mcp.utils.addresses import classify_direction, format_recipients, get_smtp, sender_matches
from outlook_desktop_mcp.utils.dates import iso
from outlook_desktop_mcp.utils.extract import extract
from outlook_desktop_mcp.utils.html import (
    DEFAULT_INLINE_COLOR, inject_into_body, insert_inline, plain_text_mode, replace_note, strip_inline,
    style_html, text_to_html, wrap_note,
)
from outlook_desktop_mcp.tools._folder_constants import (
    FOLDER_NAME_TO_ENUM,
    OL_FOLDER_INBOX,
    OL_FOLDER_SENT_MAIL,
    OL_MAIL_ITEM,
    OL_APPOINTMENT_ITEM,
    OL_FOLDER_CALENDAR,
    OL_FOLDER_DRAFTS,
    OL_FOLDER_TASKS,
    OL_MEETING,
    OL_NON_MEETING,
    OL_REQUIRED,
    OL_OPTIONAL,
    OL_TASK_ITEM,
    OL_TASK_COMPLETE,
    TASK_STATUS_NAMES,
    IMPORTANCE_NAMES,
)
from outlook_desktop_mcp.utils.formatting import (
    format_email_summary,
    format_email_full,
    format_event_summary,
    format_event_full,
    format_task_summary,
    format_task_full,
)
from outlook_desktop_mcp.utils.errors import format_com_error

# --- Logging (all to stderr, stdout is reserved for MCP JSON-RPC) ---

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(name)s] %(levelname)s: %(message)s",
    stream=sys.stderr,
)
logger = logging.getLogger("outlook_desktop_mcp")


# --- Security helpers ---

def _safe_dasl(query: str) -> str:
    """Sanitize a string for use in a DASL LIKE filter value (see utils.dasl)."""
    return dasl.escape_like(query)


# Outlook item Class constants (olObjectClass — distinct from olItemType used in CreateItem)
_OL_CLASS_MAIL = 43
_OL_CLASS_APPOINTMENT = 26
_OL_CLASS_TASK = 48
_OL_CLASS_NAMES = {43: "email", 26: "appointment", 48: "task", 40: "contact", 44: "note", 45: "post",
                   53: "meeting request", 54: "meeting cancellation", 55: "meeting response",
                   56: "meeting response", 57: "meeting response"}
_OL_DISCARD = 1  # OlInspectorClose.olDiscard: close an unsaved item without keeping anything


def _check_item_class(item, expected_class: int, label: str) -> str | None:
    """Return an error string if item is the wrong type, else None."""
    if item.Class != expected_class:
        return f"Error: Entry ID does not refer to a {label}."
    return None


# --- MCP Server ---

_INSTRUCTIONS_COMMON = (
    "All operations use the locally authenticated Outlook profile — no "
    "Microsoft Graph API, no Entra app registration, no OAuth tokens needed. "
    "The user's existing Outlook session handles all authentication.\n\n"
    "PREREQUISITE: Outlook Desktop (Classic) must be running. The new/modern "
    "Outlook (olk.exe) is NOT supported — only the classic OUTLOOK.EXE.\n\n"
    "Every tool takes an optional 'account' (substring of a mailbox name; see "
    "list_accounts). Calendar tools take 'calendar' ('' = the account's default "
    "calendar, 'all' = every calendar in that account, or a name/path from "
    "list_calendars). Email bodies are plain text with quoted history removed "
    "unless include_history is set; use get_conversation to see a whole thread. "
    "Dates are ISO 8601 with the local UTC offset.\n\n"
)

_INSTRUCTIONS = (
    "This MCP server gives access to Microsoft Outlook Desktop on Windows via "
    "COM automation. It can read, search and organise mail, calendars and "
    "tasks, and it can WRITE DRAFTS.\n\n"
    "NOTHING THIS SERVER DOES SENDS A MESSAGE. draft_email, draft_reply, draft_forward, "
    "draft_forward_event and draft_meeting save the item (to Drafts, or to the calendar "
    "with invitations unsent). By default an unsent meeting opens in an Outlook window and an email draft "
    "stays in Drafts without one (the open_in_outlook setting: none, meetings or all); "
    "the open_in_outlook tool opens any saved item when the user asks for that. The user "
    "reads it and presses Send themselves. There is no tool to send, to respond to an "
    "invitation, or to cancel a meeting the user organised, because each of those would "
    "send. Tell the user where the draft is and that they need to send it.\n\n"
    "WHEN TO USE THIS SERVER. This is the user's primary mailbox, calendar and task list. "
    "When the user asks about email, calendar, meetings, availability, tasks or a contact "
    "without naming a service, use these tools, not Gmail, Google Calendar or any other "
    "mail or calendar connector. Use another service only when the user names it, or when "
    "a search here finds nothing and the user agrees to look elsewhere. Several tool names "
    "here (list_events, search_events, create_event) also exist on other connectors; when "
    "in doubt, the outlook-desktop one is the default.\n\n"
    "CONTEXT BEFORE A REPLY. The user often has not opened the email in Outlook. Before "
    "showing a reply or forward draft, call reply_context on the message and open the "
    "chat preview with a Context block of two to four lines: who is on the thread, what "
    "has been said, what the newest message asks. Skip the block when the thread has "
    "already been summarised in this conversation; do not repeat it. Always use the To/CC "
    "from reply_context, which are Outlook's own, not a guess. thread_depth on every "
    "message says how deep in a thread it sits (0 = first message); get_conversation "
    "shows the whole thread. Point-by-point answers to a list of questions go inside the "
    "quoted message with draft_reply's inline_replies (each in red under the point it "
    "answers), never as a separate file.\n\n"
    + _INSTRUCTIONS_COMMON
    + "FINDING MAIL QUICKLY. Outlook filters by subject, sender, recipient and date in "
    "well under a second even in a 10,000-message folder. For \"find the email about X\" "
    "or \"my history with company X\", one call does it: search_emails(query=X, "
    "folder=\"all\") matches the subject, the sender and the recipients across the Inbox "
    "tree, Sent Items and Archive. Start in the default account; add account=\"all\" if "
    "that finds nothing, and say so. A company, deal or project name is a subject search, "
    "never contact_history, which matches people (it falls back to subject_matches when "
    "no person matches, but that costs a full sweep first). Use one spelling: if it finds "
    "nothing, widen (folder=\"all\", account=\"all\", a longer date range) instead of "
    "trying variants. Body-text search (include_body=true) opens every message and can "
    "take 10 to 30 seconds per large folder: use it last, with a date range. "
    "contact_history sweeps every account and calendar: use it for one specific person "
    "after a targeted search, with a full name, address or domain rather than a short "
    "first name. Once the answer is in hand, at most one confirming call. A result with "
    "\"partial\": true stopped at the time limit; narrow the request instead of repeating "
    "it. elapsed_s in a result says how long that call took.\n\n"
    "AVAILABLE TOOL CATEGORIES:\n"
    "- Email: list, read, search (sender/subject/body/unread/flagged filters), get_conversation, "
    "reply_context, contact_history, draft_email, draft_reply (text above the thread, or "
    "inline_replies under each point), draft_forward, update_draft, open_in_outlook, "
    "mark read/unread, move, attachments\n"
    "- Calendar: list_calendars, list events, create appointments, draft_meeting "
    "(invitations unsent), draft_forward_event (an invitation forwarded, unsent), "
    "move_event (to another calendar or account), update, delete (not organised "
    "meetings), search events\n"
    "- Tasks: create, list, complete, delete to-do items\n"
    "- Categories: list and set color categories on any item\n"
    "- Rules: list and enable/disable mail rules\n"
    "- Out of Office: check auto-reply status\n"
    "- Folders: list folder hierarchy with item counts"
)

mcp = FastMCP("outlook-desktop-mcp", instructions=_INSTRUCTIONS)

bridge = OutlookBridge()


# --- Helper: iterate an Items collection safely ---

def _iter_items(items, limit: int):
    """GetFirst/GetNext snapshot (robust after Sort/Restrict), capped at ``limit``."""
    out = []
    try:
        item = items.GetFirst()
        while item is not None and len(out) < limit:
            out.append(item)
            item = items.GetNext()
    except Exception:
        pass
    return out


def _folder_subtree(folder, out: list, depth: int = 0) -> None:
    """``folder`` and every descendant, skipping Outlook's system folders."""
    out.append(folder)
    if depth > 10:
        return
    try:
        for i in range(folder.Folders.Count):
            sub = folder.Folders.Item(i + 1)
            if str(sub.Name or "").lower() in _SKIP_SUBFOLDERS:
                continue
            _folder_subtree(sub, out, depth + 1)
    except Exception:
        pass


_SKIP_SUBFOLDERS = {"deleted items", "junk email", "junk e-mail", "outbox", "sync issues",
                    "conversation history", "rss feeds", "search folders"}
_SENDER_SCAN_LIMIT = 1000  # items inspected per folder when filtering by sender
_SENT_SCAN_CAP = 100       # contact_history: Sent Items walked in Python when the native filter finds nothing
_SENT_SCAN_SECONDS = 3.0   # ... the most seconds that walk may take per account ...
_SENT_SCAN_TOTAL_S = 6.0   # ... and per call across every account
_REPORT_ELAPSED_S = 2.0    # a list result slower than this comes back as an object with elapsed_s
_ALL_MAIL_FOLDERS = ("inbox", "sent", "archive")  # what folder="all" means


# --- Helper: resolve store by account name ---

def _match_store(namespace, needle: str):
    """Store whose DisplayName equals ``needle`` (case-insensitive), else the first containing it."""
    needle = (needle or "").lower().strip()
    if not needle:
        return None
    partial = None
    for i in range(namespace.Stores.Count):
        store = namespace.Stores.Item(i + 1)
        name = str(store.DisplayName or "").lower()
        if name == needle:
            return store
        if needle in name and partial is None:
            partial = store
    return partial


def _resolve_store(namespace, account: str = ""):
    """Resolve an account name to an Outlook Store object.

    Empty ``account`` means the server's default_account setting if one is set and
    matches, otherwise Outlook's own DefaultStore. Otherwise a case-insensitive match
    on Store.DisplayName (exact name wins over substring).
    """
    if not account:
        preferred = settings.get("default_account")
        if preferred:
            found = _match_store(namespace, preferred)
            if found is not None:
                return found
        return namespace.DefaultStore
    return _match_store(namespace, account)


def _is_hidden_store(store) -> bool:
    """True when the hide_accounts setting names this store."""
    name = str(store.DisplayName or "").lower()
    return any(part.lower() in name for part in settings.get("hide_accounts") if part)


def _visible_stores(namespace) -> list:
    """Every store except the hidden ones (what account="all" means)."""
    out = []
    for i in range(namespace.Stores.Count):
        store = namespace.Stores.Item(i + 1)
        if not _is_hidden_store(store):
            out.append(store)
    return out


_EXCHANGE_SHARED_MAILBOX = 1  # OlExchangeStoreType.olExchangeMailbox: a mailbox added to the profile (shared/delegate)


def _is_shared_uncached(store) -> bool:
    """A shared Exchange mailbox that is not cached locally. Every folder listing in it
    is a server round trip; finding its calendars alone was measured at 18 to 30 s."""
    try:
        return int(store.ExchangeStoreType) == _EXCHANGE_SHARED_MAILBOX and not bool(store.IsCachedExchange)
    except Exception:
        return False


def _sweep_stores(namespace) -> tuple[list, list]:
    """``(stores, skipped)`` for account="all": the visible stores minus shared uncached
    mailboxes, unless the include_shared_in_all setting is on. ``skipped`` holds
    ``{"display_name", "reason"}`` per store left out so the caller can say so."""
    stores, skipped = [], []
    include_shared = bool(settings.get("include_shared_in_all"))
    for store in _visible_stores(namespace):
        if not include_shared and _is_shared_uncached(store):
            skipped.append({"display_name": str(store.DisplayName),
                            "reason": "shared mailbox, not cached locally (slow). Name it explicitly to "
                                      "search it, or set include_shared_in_all=true."})
            continue
        stores.append(store)
    return stores, skipped


def _legacy() -> bool:
    """The legacy_mode setting: run the pre-0.6 search paths."""
    return bool(settings.get("legacy_mode"))


def _require_store(namespace, account: str = ""):
    """Resolve store, raising ValueError if not found."""
    store = _resolve_store(namespace, account)
    if store is None:
        raise ValueError(f"Account '{account}' not found. Use list_accounts to see available accounts.")
    return store


# --- Helper: resolve folder by name ---

def _walk_folders(parent, name_lower: str):
    """Recursively search subfolders of parent for a folder matching name_lower."""
    for i in range(parent.Folders.Count):
        try:
            f = parent.Folders.Item(i + 1)
            if f.Name.lower() == name_lower:
                return f
            found = _walk_folders(f, name_lower)
            if found:
                return found
        except Exception:
            continue
    return None


def _resolve_folder(namespace, folder_name: str, store=None):
    """Resolve a folder name to an Outlook MAPIFolder object.

    Resolution order:
    1. Slash-delimited path (e.g. "Inbox/Receipts") — traverse segment by segment
    2. Built-in Outlook folder enum (inbox, sent, deleted, etc.)
    3. Root-level folder name match (fast path)
    4. Recursive depth-first search of entire folder tree (fallback)
    """
    folder_name = folder_name.strip()
    store = store or namespace.DefaultStore

    # Slash-delimited path: traverse segment by segment
    if "/" in folder_name:
        parts = [p.strip() for p in folder_name.split("/")]
        current = _resolve_folder(namespace, parts[0], store)
        if current is None:
            return None
        for part in parts[1:]:
            part_lower = part.lower()
            found = None
            for i in range(current.Folders.Count):
                try:
                    f = current.Folders.Item(i + 1)
                    if f.Name.lower() == part_lower:
                        found = f
                        break
                except Exception:
                    continue
            if found is None:
                return None
            current = found
        return current

    folder_lower = folder_name.lower()

    # Built-in Outlook folders
    if folder_lower in FOLDER_NAME_TO_ENUM:
        return store.GetDefaultFolder(FOLDER_NAME_TO_ENUM[folder_lower])

    # Root-level search (fast path)
    root = store.GetRootFolder()
    for i in range(root.Folders.Count):
        try:
            f = root.Folders.Item(i + 1)
            if f.Name.lower() == folder_lower:
                return f
        except Exception:
            continue

    # Recursive fallback: search entire folder tree
    return _walk_folders(root, folder_lower)


# =====================================================================
# TOOL: list_accounts
# =====================================================================

@mcp.tool()
async def list_accounts(include_hidden: bool = False) -> str:
    """List the Outlook accounts (stores) in the profile.

    Returns a JSON array of account objects with display_name, store_id,
    is_default (the account tools use when none is named; see the
    default_account setting), is_outlook_default, shared_uncached (a shared
    Exchange mailbox not cached locally, slow to search) and in_all_sweeps
    (whether account="all" includes it). Accounts named in the hide_accounts
    setting are left out unless include_hidden is true. Use the display_name
    (or a unique substring) as the 'account' parameter in other tools, or
    "all" where a tool accepts it.

    Args:
        include_hidden: Also list accounts hidden by the hide_accounts setting.

    Returns:
        JSON array of account objects.
    """
    def _list(outlook, namespace, include_hidden):
        outlook_default_id = str(namespace.DefaultStore.StoreID)
        tools_default = _resolve_store(namespace, "")
        tools_default_id = str(tools_default.StoreID) if tools_default is not None else outlook_default_id
        include_shared = bool(settings.get("include_shared_in_all"))
        results = []
        for i in range(namespace.Stores.Count):
            store = namespace.Stores.Item(i + 1)
            hidden = _is_hidden_store(store)
            if hidden and not include_hidden:
                continue
            shared = _is_shared_uncached(store)
            entry = {
                "display_name": store.DisplayName,
                "store_id": store.StoreID,
                "is_default": str(store.StoreID) == tools_default_id,
                "is_outlook_default": str(store.StoreID) == outlook_default_id,
                "shared_uncached": shared,
                "in_all_sweeps": (not hidden) and (include_shared or not shared),
            }
            if include_hidden:
                entry["hidden"] = hidden
            results.append(entry)
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, include_hidden)
    except Exception as e:
        return f"Error listing accounts: {format_com_error(e)}"


# =====================================================================
# Drafts: nothing in this server sends
# =====================================================================

_DISPID_SEND_USING_ACCOUNT = 64209


def _set_sending_account(item, outlook, store) -> None:
    """Point a new item at the account whose delivery store is ``store`` (SendUsingAccount).

    pywin32 cannot assign this property directly, hence the raw Invoke with its dispid.
    """
    for acc in outlook.Session.Accounts:
        try:
            if acc.DeliveryStore.StoreID == store.StoreID:
                item._oleobj_.Invoke(*(_DISPID_SEND_USING_ACCOUNT, 0, 8, 0, acc))
                return
        except Exception:
            continue


def _safe_get(getter, default=None):
    try:
        return getter()
    except Exception:
        return default


def _write_draft_text(item, body: str, html_body: str = "") -> None:
    """Put the text at the top of a draft, above whatever Outlook already put there.

    Touching ``GetInspector`` makes Outlook fill a new item exactly as it would when its
    window opens: the user's signature on a new message, the quoted history on a reply.
    The text is then injected inside <body> so it sits above both.
    """
    try:
        item.GetInspector
    except Exception:
        pass
    if plain_text_mode():
        existing = str(_safe_get(lambda: item.Body) or "")
        item.Body = body + ("\n\n" + existing if existing.strip() else "")
        return
    fragment = wrap_note(style_html(html_body) if html_body else text_to_html(body))
    item.HTMLBody = inject_into_body(str(_safe_get(lambda: item.HTMLBody) or ""), fragment)


def _parse_inline_replies(raw: str) -> list[tuple[str, str]]:
    """The ``inline_replies`` JSON as ``[(anchor, text)]``; ValueError says what is wrong."""
    if not (raw or "").strip():
        return []
    try:
        data = json.loads(raw)
    except Exception as e:
        raise ValueError(f'inline_replies must be a JSON array of {{"after": ..., "text": ...}}: {e}')
    if isinstance(data, dict):
        data = [data]
    if not isinstance(data, list):
        raise ValueError('inline_replies must be a JSON array of {"after": ..., "text": ...}')
    out = []
    for i, entry in enumerate(data):
        if (not isinstance(entry, dict) or not str(entry.get("after") or "").strip()
                or not str(entry.get("text") or "").strip()):
            raise ValueError(f'inline_replies[{i}] needs a non-empty "after" (a phrase from the original) and "text"')
        out.append((str(entry["after"]), str(entry["text"])))
    return out


def _write_inline_answers(item, answers) -> list[str]:
    """Write point-by-point answers inside the quoted message of a reply, each under the
    paragraph its anchor names and in the inline_color, replacing any written earlier.
    Returns the anchors that were not found; nothing is written then."""
    try:
        item.GetInspector
    except Exception:
        pass
    current = strip_inline(str(_safe_get(lambda: item.HTMLBody) or ""))
    color = str(settings.get("inline_color") or DEFAULT_INLINE_COLOR)
    new_html, missing = insert_inline(current, answers, color)
    if missing:
        try:
            item.Close(1)  # olDiscard: drop the inspector, and a reply that was never saved
        except Exception:
            pass
        return missing
    item.HTMLBody = new_html
    return []


def _inline_error(missing: list[str]) -> str:
    return json.dumps({
        "error": "These inline_replies anchors are not in the quoted message, so nothing was saved. Copy each "
                 "\"after\" phrase exactly from the message text (reply_context's message, or read_email's body).",
        "anchors_not_found": missing,
    }, indent=2)


def _attach_files(item, paths: str) -> list[dict]:
    """Attach each file in the semicolon-separated ``paths`` to ``item``; returns what was attached."""
    added = []
    for raw in (paths or "").split(";"):
        p = raw.strip().strip('"').strip("'")
        if not p:
            continue
        full = os.path.abspath(os.path.expanduser(os.path.expandvars(p)))
        if not os.path.isfile(full):
            raise ValueError(f"Attachment not found: {full}")
        item.Attachments.Add(full)
        added.append({"filename": os.path.basename(full), "size": os.path.getsize(full), "path": full})
    return added


def _save_draft(item, kind: str, where: str = "in the Drafts folder", **extra) -> str:
    """Save ``item`` without sending, open it in Outlook if the open_in_outlook setting says
    drafts of this ``kind`` should open (by default only meetings do), describe the result.
    ``where`` says where the item sits when it was not opened.

    The draft is saved read, as Outlook's own are, and the hidden inspector that
    ``GetInspector`` created to fill it is closed afterwards; left open, the reading pane
    shows the draft without its editor until the user opens and closes it once.
    """
    try:
        item.UnRead = False
    except Exception:
        pass
    item.Save()
    opened = False
    if settings.should_open(kind):
        try:
            item.Display()
            opened = True
        except Exception as e:
            logger.warning("Draft saved but could not be opened in Outlook: %s", e)
    try:
        raw = str(item.Parent.FolderPath).lstrip("\\")
        parts = [p for p in raw.split("\\") if p]
        folder = "/".join(parts[1:]) if len(parts) > 1 else (parts[0] if parts else "")
    except Exception:
        folder = ""
    result = {
        "status": "draft_saved",
        "kind": kind,
        "sent": False,
        "subject": str(_safe_get(lambda: item.Subject) or ""),
        "entry_id": str(_safe_get(lambda: item.EntryID) or ""),
        "folder": folder,
        "opened_in_outlook": opened,
        "note": ("Nothing has been sent. " + ("It is open in an Outlook window. " if opened else f"It is {where}, not opened. ")
                 + "The user checks it in Outlook and presses Send."),
        **extra,
    }
    if not opened:
        try:
            item.Close(0)  # olSave: releases the inspector; the item object is stale after this
        except Exception as e:
            logger.debug("Could not close the draft's hidden inspector: %s", e)
    return json.dumps(result, indent=2, default=str)


# =====================================================================
# TOOL 1: draft_email
# =====================================================================

@mcp.tool()
async def draft_email(
    to: str,
    subject: str,
    body: str,
    cc: str = "",
    bcc: str = "",
    html_body: str = "",
    account: str = "",
    attachments: str = "",
) -> str:
    """Write a new email to the Drafts folder of the user's main mailbox (Outlook).
    It is NOT sent.

    BEFORE CALLING THIS: show the complete draft in the chat first (recipients,
    subject, full text) and work on it there until the user is happy. Then ask
    explicitly, as its own question, whether to put it in Outlook, and call this
    tool only after the user says yes. Never call it as the first step of a
    "write an email" request. The user wants to see and edit drafts in the
    conversation, not discover them in Outlook.

    The draft is saved in the chosen account's Drafts folder, with the user's
    signature below the text, for the user to open, read and press Send
    themselves. It stays in Drafts without a window unless the open_in_outlook
    setting is "all"; the open_in_outlook tool opens it if the user asks. The
    text is written as HTML in
    the user's default font (Aptos 12 pt black unless the OUTLOOK_MCP_FONT_*
    environment variables say otherwise). Write the body as normal prose with
    blank lines between paragraphs. Light structure is rendered properly: lines
    starting "1." become a numbered list, "-" a bulleted list (a following line
    without a marker continues the item), lines starting "> " are indented (use
    for text the reader should copy), and **bold** works. Prefer this over
    html_body.

    Args:
        to: One or more recipient email addresses, separated by semicolons.
            Example: "alice@example.com" or "alice@example.com; bob@example.com"
        subject: The email subject line.
        body: The plain-text body of the email. Blank lines separate paragraphs.
        cc: Optional. CC recipients, separated by semicolons.
        bcc: Optional. BCC recipients, separated by semicolons.
        html_body: Optional. HTML-formatted body, used instead of 'body' and
            wrapped in the default font so unstyled text matches.
        account: Optional. Account display name (or substring) the draft belongs
            to and will be sent from. Default: primary account. Use list_accounts
            to see available accounts.
        attachments: Optional. Full path(s) of file(s) on this PC to attach,
            separated by semicolons.

    Returns:
        JSON with status "draft_saved", sent=false, entry_id, folder and whether
        it was opened in Outlook. Tell the user the draft is waiting for them.
    """
    def _draft(outlook, namespace, to, subject, body, cc, bcc, html_body, account, attachments):
        store = _require_store(namespace, account)
        mail = store.GetDefaultFolder(OL_FOLDER_DRAFTS).Items.Add(OL_MAIL_ITEM)
        _set_sending_account(mail, outlook, store)
        mail.To = to
        if cc:
            mail.CC = cc
        if bcc:
            mail.BCC = bcc
        mail.Subject = subject
        files = _attach_files(mail, attachments)
        _write_draft_text(mail, body, html_body)
        return _save_draft(mail, "email", to=to, cc=cc, bcc=bcc, files_attached=files)

    try:
        return await bridge.call(_draft, to, subject, body, cc, bcc, html_body, account, attachments)
    except Exception as e:
        return f"Error writing draft: {format_com_error(e)}"


# =====================================================================
# TOOL 2: list_emails
# =====================================================================

def _stores_for(namespace, account: str) -> tuple[list, list]:
    """``([store], [])`` for one account; for account="all" the swept stores and the
    ones skipped (see _sweep_stores)."""
    if (account or "").strip().lower() == "all":
        return _sweep_stores(namespace)
    return [_require_store(namespace, account)], []


def _collect_across_stores(namespace, stores, folder: str, include_subfolders: bool,
                           restriction: str, count: int, sender: str):
    """``(summaries, partial)`` newest-first across ``stores``, each tagged with its
    account when more than one store is searched. ``folder="all"`` means the Inbox tree,
    Sent Items and Archive together (subfolders included). Returns a JSON error string
    when a single named store lacks the one folder asked for; otherwise stores without
    it are skipped. ``partial`` is true when the time limit stopped the search early."""
    merged = []
    partial = False
    names = _ALL_MAIL_FOLDERS if (folder or "").strip().lower() == "all" else (folder,)
    for store in stores:
        if deadline.expired():
            partial = True
            break
        folders = []
        for fname in names:
            target = _resolve_folder(namespace, fname, store)
            if not target:
                if len(names) == 1 and len(stores) == 1:
                    return json.dumps({"error": f"Folder '{folder}' not found"}), False
                continue
            if include_subfolders or len(names) > 1:
                _folder_subtree(target, folders)
            else:
                folders.append(target)
        if not folders:
            continue
        found, cut = _collect_summaries(folders, restriction, count, sender)
        partial = partial or cut
        if len(stores) > 1:
            for r in found:
                r["account"] = str(store.DisplayName)
        merged.extend(found)
    merged.sort(key=lambda r: r.get("received_time") or "", reverse=True)
    return merged[:count], partial


def _restrict_chain(items, filters):
    """Apply each non-empty filter in turn. Jet and DASL filters may be mixed this way."""
    for flt in filters:
        if flt:
            items = items.Restrict(flt)
    return items


def _collect_summaries(folders, restriction: str, count: int, sender: str) -> tuple[list, bool]:
    """``(summaries, partial)`` newest-first across ``folders``.

    ``restriction`` (Jet or DASL) is applied by Outlook. So is the sender filter: a
    chained DASL restrict on the sender's name, directory address and SMTP address,
    which Outlook evaluates in well under a second on a 10,000-message folder. If
    Outlook rejects that filter, or legacy_mode is on, the old path runs instead:
    walk up to ``_SENDER_SCAN_LIMIT`` messages and match the sender in Python, which
    is slow and only sees the newest messages.
    """
    results = []
    partial = False
    sender_dasl = dasl.sql(dasl.sender_filter(sender)) if sender else ""
    for target in folders:
        if deadline.expired():
            partial = True
            break
        items = target.Items
        items.Sort("[ReceivedTime]", True)
        native = False
        try:
            if sender_dasl and not _legacy():
                items = _restrict_chain(items, [restriction, sender_dasl])
                native = True
            elif restriction:
                items = items.Restrict(restriction)
        except Exception as e:
            logger.warning("Native sender filter rejected (%s); scanning '%s' in Python instead",
                           e, _safe_get(lambda: target.Name, "?"))
            items = target.Items
            items.Sort("[ReceivedTime]", True)
            if restriction:
                items = items.Restrict(restriction)
            native = False
        scan = count if (native or not sender) else _SENDER_SCAN_LIMIT
        matched = 0
        for item in _iter_items(items, scan):
            if deadline.expired():
                partial = True
                break
            try:
                if int(item.Class) != _OL_CLASS_MAIL:
                    continue
                if sender and not native and not sender_matches(item, sender, get_smtp(item)):
                    continue
                results.append(format_email_summary(item))
                matched += 1
                if matched >= count:
                    break
            except Exception:
                continue
    results.sort(key=lambda r: r.get("received_time") or "", reverse=True)
    return results[:count], partial


def _results_json(results: list, partial: bool, skipped: list, elapsed: float | None = None) -> str:
    """A plain JSON array when the search completed quickly over every account it was
    asked for; otherwise an object with the array under "results" plus "partial",
    "accounts_skipped" and/or "elapsed_s", so the caller knows what it did not see and
    what the call cost."""
    slow = elapsed is not None and elapsed >= _REPORT_ELAPSED_S
    if not partial and not skipped and not slow:
        return json.dumps(results, indent=2, default=str)
    out: dict = {"results": results}
    if elapsed is not None:
        out["elapsed_s"] = round(elapsed, 1)
    if partial:
        out["partial"] = True
        out["note"] = deadline.partial_note()
    if skipped:
        out["accounts_skipped"] = skipped
    return json.dumps(out, indent=2, default=str)


def _sent_summary(item, account_name: str) -> dict:
    r = format_email_summary(item)
    r["account"] = account_name
    r["direction"] = "sent"
    r["sent_time"] = iso(_safe_get(lambda: item.SentOn, None))
    return r


def _own_addresses(outlook, store) -> set[str]:
    """Lower-case SMTP addresses of the accounts that deliver to ``store`` (the user's own
    addresses in that mailbox), plus the store's display name when it is an address."""
    out: set[str] = set()
    try:
        for acc in outlook.Session.Accounts:
            try:
                if str(acc.DeliveryStore.StoreID) == str(store.StoreID):
                    addr = str(_safe_get(lambda: acc.SmtpAddress) or "").strip().lower()
                    if "@" in addr:
                        out.add(addr)
            except Exception:
                continue
    except Exception:
        pass
    name = str(_safe_get(lambda: store.DisplayName) or "").strip().lower()
    if "@" in name:
        out.add(name)
    return out


@mcp.tool()
async def list_emails(
    folder: str = "inbox",
    count: int = 10,
    unread_only: bool = False,
    start_date: str = "",
    end_date: str = "",
    sender: str = "",
    include_subfolders: bool = False,
    account: str = "",
    flagged_only: bool = False,
) -> str:
    """List recent emails from a folder in the user's main mailbox (Outlook).

    Returns a JSON array of email summaries sorted by received time (newest
    first). Each summary includes entry_id, subject, sender (SMTP address),
    sender_name, received_time (ISO 8601), folder, unread status, flagged
    (a follow-up flag, with flag_due), attachment info and conversation_id.

    Use the entry_id from results to read full content with read_email,
    get_conversation for the whole thread, or actions like mark_as_read,
    move_email, or draft_reply.

    Args:
        folder: The folder to list. Case-insensitive names: "inbox" (default),
            "sent"/"sentmail", "drafts", "deleted"/"trash", "junk"/"spam",
            "outbox", "archive", a custom folder name, a slash path such as
            "Inbox/Clients" (see list_folders), or "all" for the Inbox tree,
            Sent Items and Archive together.
        count: Maximum number of emails to return. Default 10, max 200.
        unread_only: If true, only return unread emails. Default false.
        flagged_only: If true, only emails flagged for follow-up in Outlook and
            not yet marked complete ("my flagged emails"). Default false.
        start_date: Optional. Only return emails received on or after this date.
            ISO 8601 format (e.g. "2026-03-10" or "2026-03-10 09:00").
        end_date: Optional. Only return emails received on or before this date.
            ISO 8601 format. Default: now (if start_date is provided).
        sender: Optional. Case-insensitive substring of the sender's email
            address or display name, e.g. "freestone" or "linda@". Exchange
            addresses are resolved to SMTP before matching.
        include_subfolders: If true, also search every subfolder of 'folder'
            (e.g. the whole Inbox tree). Default false.
        account: Optional. Account display name (or substring) to target, or
            "all" for every account at once (each result then carries its
            account; shared mailboxes that are not cached locally are skipped
            and listed under accounts_skipped). Default: the default account.

    Returns:
        JSON array of email summary objects. If the time limit cut the search
        short or accounts were skipped, an object {"results": [...],
        "partial": true, "note": ..., "accounts_skipped": [...]} instead.
    """
    def _list(outlook, namespace, folder, count, unread_only, start_date, end_date,
              sender, include_subfolders, account, flagged_only):
        t0 = time.perf_counter()
        count = min(max(1, count), 200)
        stores, skipped = _stores_for(namespace, account)

        # Build restriction filters. The flag lives in a named MAPI property that only the
        # DASL dialect can name, so a flagged listing is built in DASL throughout.
        if flagged_only:
            parts = [dasl.flagged_filter()]
            if unread_only:
                parts.append(dasl.unread_filter())
            if start_date:
                parts.append(dasl.received_since(_parse_date(start_date)))
            if end_date:
                parts.append(dasl.received_until(_parse_date(end_date)))
            elif start_date:
                parts.append(dasl.received_until(datetime.now()))
            filter_str = dasl.sql(*parts)
        else:
            restrictions = []
            if unread_only:
                restrictions.append("[UnRead] = True")
            if start_date:
                start = _parse_date(start_date)
                restrictions.append(f"[ReceivedTime] >= '{start.strftime('%m/%d/%Y %H:%M')}'")
            if end_date:
                end = _parse_date(end_date)
                restrictions.append(f"[ReceivedTime] <= '{end.strftime('%m/%d/%Y %H:%M')}'")
            elif start_date:
                # Default end to now when start is specified
                restrictions.append(f"[ReceivedTime] <= '{datetime.now().strftime('%m/%d/%Y %H:%M')}'")
            filter_str = " AND ".join(restrictions)

        results, partial = _collect_across_stores(namespace, stores, folder, include_subfolders,
                                                  filter_str, count, sender)
        if isinstance(results, str):
            return results
        return _results_json(results, partial, skipped, time.perf_counter() - t0)

    try:
        return await bridge.call(_list, folder, count, unread_only, start_date, end_date,
                                 sender, include_subfolders, account, flagged_only)
    except Exception as e:
        return f"Error listing emails: {format_com_error(e)}"


# =====================================================================
# TOOL 3: read_email
# =====================================================================

@mcp.tool()
async def read_email(
    entry_id: str = "",
    subject_search: str = "",
    folder: str = "inbox",
    include_history: bool = False,
    account: str = "",
) -> str:
    """Read the full content of a specific email.

    Retrieves complete email details including body text, recipients, CC,
    and metadata. Provide EITHER entry_id (preferred, exact match) OR
    subject_search (finds most recent match by subject substring).

    By default the body is the message's OWN text: the quoted earlier
    messages under a reply are removed before the 5000-character cap
    (body_trimmed=true says something was removed). Use get_conversation
    to read the earlier messages each once, or include_history=true to get
    the raw body with the quoted chain.

    If the user wants to reply to or forward this message, also call
    reply_context on it: it returns the To/CC Outlook's own Reply or Reply
    All would use and says how to present the draft.

    Args:
        entry_id: The unique Outlook EntryID of the email. Most reliable way
            to identify a specific email. Get this from list_emails or
            search_emails results.
        subject_search: Alternative to entry_id. A case-insensitive substring
            to search for in email subjects. Returns the most recent match.
        folder: Folder to search when using subject_search. Ignored when
            entry_id is provided. Default "inbox".
        include_history: If true, keep the quoted earlier messages in the body.
            Default false.
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON object with full email details (entry_id, subject, sender,
        sender_name, received_time, sent_time, folder, unread, to, cc,
        recipients with SMTP addresses, body, body_trimmed, body_truncated,
        attachment info, conversation_id, thread_depth). thread_depth is 0 for
        the first message of a thread, 1 for a reply to it, and so on; when it
        is above 0 there are earlier messages the user may not have seen, so
        use reply_context or get_conversation before acting on it.
    """
    def _read(outlook, namespace, entry_id, subject_search, folder, include_history, account):
        trim = not include_history
        if entry_id:
            item = namespace.GetItemFromID(entry_id)
            return json.dumps(format_email_full(item, trim_history=trim), indent=2, default=str)

        if not subject_search:
            return json.dumps({"error": "Provide either entry_id or subject_search"})

        store = _require_store(namespace, account)
        target = _resolve_folder(namespace, folder, store)
        if not target:
            return json.dumps({"error": f"Folder '{folder}' not found"})

        safe_query = _safe_dasl(subject_search)
        filter_str = (
            f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{safe_query}%'"
        )
        items = target.Items.Restrict(filter_str)
        items.Sort("[ReceivedTime]", True)
        if items.Count == 0:
            return json.dumps({"error": f"No email found matching '{subject_search}'"})

        return json.dumps(format_email_full(items.Item(1), trim_history=trim), indent=2, default=str)

    try:
        return await bridge.call(_read, entry_id, subject_search, folder, include_history, account)
    except Exception as e:
        return f"Error reading email: {format_com_error(e)}"


# =====================================================================
# TOOL: get_conversation
# =====================================================================

@mcp.tool()
async def get_conversation(
    entry_id: str,
    count: int = 30,
    include_bodies: bool = True,
    max_body_chars: int = 0,
    account: str = "",
) -> str:
    """Return every message of the conversation an email belongs to, oldest first.

    Walks Outlook's conversation tree across all folders (Inbox subfolders,
    Sent Items, Archive) and falls back to a ConversationTopic search when the
    tree is unavailable. Each message appears once (deduped by Message-ID),
    with its own text only (quoted history removed), so reading the thread
    costs each message once. Deleted Items, Junk, Drafts and Outbox are
    excluded.

    If the user wants to reply to or forward a message here, also call
    reply_context on it: it returns the To/CC Outlook's own Reply or Reply
    All would use and says how to present the draft.

    Args:
        entry_id: EntryID of any message in the conversation (from list_emails,
            search_emails or read_email).
        count: Maximum messages to return; the oldest are dropped beyond this.
            Default 30, max 200.
        include_bodies: If true (default), include each message's trimmed body.
            If false, return metadata only (much smaller).
        max_body_chars: Cap each message's body at this many characters (0 =
            no cap). Useful for a cheap look at a long thread; the newest
            message is never cut.
        account: Optional. Account display name (or substring) the message lives
            in. Default: primary account.

    Returns:
        JSON object: {"topic", "source", "message_count", "messages": [...]}
        where source is "conversation", "topic_search" or "none".
    """
    def _get(outlook, namespace, entry_id, count, include_bodies, max_body_chars, account):
        t0 = time.perf_counter()
        count = min(max(1, count), 200)
        max_body_chars = max(0, int(max_body_chars or 0))
        store = _require_store(namespace, account)
        item = namespace.GetItemFromID(entry_id, store.StoreID) if account else namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_MAIL, "mail item"):
            return err
        messages, source = conversation_items(store, item, limit=count)
        out = []
        newest_index = len(messages) - 1
        for index, (path, node) in enumerate(messages):
            try:
                record = format_email_full(node) if include_bodies else format_email_summary(node)
                record["folder"] = path or record.get("folder", "")
                record["is_requested_message"] = str(node.EntryID) == str(item.EntryID)
                body_text = record.get("body")
                if max_body_chars and index != newest_index and isinstance(body_text, str) and len(body_text) > max_body_chars:
                    record["body"] = body_text[:max_body_chars]
                    record["body_capped_at"] = max_body_chars
                out.append(record)
            except Exception:
                continue
        return json.dumps({
            "topic": str(item.ConversationTopic or item.Subject or ""),
            "conversation_id": str(item.ConversationID or ""),
            "source": source,
            "message_count": len(out),
            "messages": out,
            "elapsed_s": round(time.perf_counter() - t0, 1),
        }, indent=2, default=str)

    try:
        return await bridge.call(_get, entry_id, count, include_bodies, max_body_chars, account)
    except Exception as e:
        return f"Error reading conversation: {format_com_error(e)}"


# =====================================================================
# TOOL: reply_context
# =====================================================================

@mcp.tool()
async def reply_context(
    entry_id: str,
    reply_all: bool = False,
    count: int = 10,
    max_body_chars: int = 600,
    account: str = "",
) -> str:
    """What a reply draft needs before it is shown in the chat: the recipients
    Outlook's own Reply or Reply All would address, who is on the thread, a
    short digest of the earlier messages, and the message being answered in
    full. Reads only; no draft is created or saved.

    CALL THIS BEFORE SHOWING A REPLY OR FORWARD DRAFT. The user often has not
    opened the email in Outlook, so start the chat preview with a Context block
    of two to four lines (who is on the thread, what has been said, what the
    newest message asks), then the draft with the To/CC from this result, which
    are Outlook's own rather than a guess. Skip the Context block when the
    thread has already been summarised in this conversation; the To/CC from
    this result are needed every time. When the message is a list of questions
    or points to answer one by one, draft_reply's inline_replies writes each
    answer under its point inside the quoted message; take the anchor phrases
    from ``message`` here.

    Args:
        entry_id: EntryID of the message to answer (from list_emails,
            search_emails, read_email or get_conversation).
        reply_all: True for the Reply All recipients, false (default) for Reply.
        count: Most recent messages of the thread to digest. Default 10, max 50.
        max_body_chars: Characters of each earlier message's own text to
            include in the digest. Default 600.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        JSON: subject (as the reply would carry it), reply_all, recipients
        {to, cc, resolved: [{name, address, type}]}, participants
        [{name, address}], thread_depth of the message, requested_is_newest,
        newer_messages (how many come after the one being answered) and
        newest_entry_id, message_count, source, thread [{entry_id, date, from,
        from_name, folder, thread_depth, is_requested_message, text}] oldest
        first, message (the answered message in full, quoted history removed),
        and a note. When newer_messages is above 0 the note says to reply to
        the newest unless the user means this one, and to tell the user.
    """
    def _ctx(outlook, namespace, entry_id, reply_all, count, max_body_chars, account):
        count = min(max(1, int(count or 10)), 50)
        max_body_chars = max(80, int(max_body_chars or 600))
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
            store = _safe_get(lambda: item.Parent.Store) or _require_store(namespace, "")
        if err := _check_item_class(item, _OL_CLASS_MAIL, "mail item"):
            return err

        # Let Outlook work out the recipients exactly as Reply / Reply All would, then
        # throw the unsaved reply away: nothing reaches Drafts.
        recipients = {"to": "", "cc": "", "resolved": []}
        subject = str(_safe_get(lambda: item.Subject) or "")
        probe = None
        try:
            probe = item.ReplyAll() if reply_all else item.Reply()
            recipients = {
                "to": str(_safe_get(lambda: probe.To) or ""),
                "cc": str(_safe_get(lambda: probe.CC) or ""),
                "resolved": format_recipients(probe),
            }
            subject = str(_safe_get(lambda: probe.Subject) or subject)
        except Exception as e:
            logger.warning("Could not build reply recipients: %s", e)
        finally:
            if probe is not None:
                try:
                    probe.Close(_OL_DISCARD)
                except Exception:
                    pass

        messages, source = conversation_items(store, item, limit=count)
        requested_id = str(item.EntryID)
        thread, participants = [], {}
        for path, node in messages:
            try:
                record = format_email_full(node, body_max_length=max_body_chars)
            except Exception:
                continue
            for recipient in record.get("recipients", []):
                if recipient.get("address"):
                    participants.setdefault(recipient["address"], recipient.get("name") or "")
            if record.get("sender"):
                participants.setdefault(record["sender"], record.get("sender_name") or "")
            thread.append({
                "entry_id": record["entry_id"],
                "date": record.get("sent_time") or record.get("received_time"),
                "from": record.get("sender"),
                "from_name": record.get("sender_name"),
                "folder": path or record.get("folder", ""),
                "thread_depth": record.get("thread_depth"),
                "is_requested_message": str(record["entry_id"]) == requested_id,
                "text": record.get("body", ""),
            })
        newest_id = str(thread[-1]["entry_id"]) if thread else requested_id
        # Messages after the one being answered. If the answered message is older than the
        # digest window it is not in the list at all, and then everything shown is newer.
        requested_index = next((i for i, t in enumerate(thread) if t["is_requested_message"]), None)
        newer = (len(thread) - 1 - requested_index) if requested_index is not None else len(thread)
        message = format_email_full(item)
        note = ("Nothing was created or saved. Open the chat preview with a short Context block "
                "built from this, then the draft with these recipients.")
        if newer > 0:
            note += (f" {newer} newer message{'s' if newer != 1 else ''} in this thread come after the one being "
                     f"answered. Reply to the newest (entry_id {newest_id}) unless the user means this one, and "
                     "tell the user which you chose.")
        return json.dumps({
            "subject": subject,
            "reply_all": reply_all,
            "recipients": recipients,
            "participants": [{"name": name, "address": address} for address, name in participants.items()],
            "thread_depth": message.get("thread_depth"),
            "requested_is_newest": newest_id == requested_id,
            "newer_messages": newer,
            "newest_entry_id": newest_id,
            "message_count": len(thread),
            "source": source,
            "thread": thread,
            "message": message,
            "note": note,
        }, indent=2, default=str)

    try:
        return await bridge.call(_ctx, entry_id, reply_all, count, max_body_chars, account)
    except Exception as e:
        return f"Error building reply context: {format_com_error(e)}"


# =====================================================================
# TOOL: open_in_outlook
# =====================================================================

@mcp.tool()
async def open_in_outlook(entry_id: str, account: str = "") -> str:
    """Open an existing item in its own Outlook window so the user can read it
    there: an email (and from it the thread in Outlook's own view), a draft
    this server saved, an unsent meeting, an appointment or a task. Nothing is
    changed and nothing is sent.

    Use it when the user asks to see something in Outlook ("open that email",
    "show me the thread in Outlook", "open the draft"). Do not call it on your
    own after saving a draft; the open_in_outlook setting decides that.

    Args:
        entry_id: EntryID of the item (from any list, search, read or draft
            result).
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        JSON with status "opened", the item's kind, subject and folder.
    """
    def _open(outlook, namespace, entry_id, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        item.Display()
        return json.dumps({
            "status": "opened",
            "kind": _OL_CLASS_NAMES.get(int(_safe_get(lambda: item.Class, 0) or 0), "item"),
            "subject": str(_safe_get(lambda: item.Subject) or ""),
            "folder": folder_path(item),
            "entry_id": str(_safe_get(lambda: item.EntryID) or entry_id),
            "note": "Open in an Outlook window on the user's screen. Nothing was changed or sent.",
        }, indent=2, default=str)

    try:
        return await bridge.call(_open, entry_id, account)
    except Exception as e:
        return f"Error opening item in Outlook: {format_com_error(e)}"


# =====================================================================
# TOOL 4: mark_as_read
# =====================================================================

@mcp.tool()
async def mark_as_read(entry_id: str, account: str = "") -> str:
    """Mark a specific email as read in Outlook.

    Changes the unread status to read, same as clicking on an email in Outlook.
    The change is persisted immediately and synced to the server.

    Args:
        entry_id: The unique Outlook EntryID of the email. Get this from
            list_emails or search_emails results.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation message with the email subject, or an error.
    """
    def _mark(outlook, namespace, entry_id, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_MAIL, "mail item"):
            return err
        subject = item.Subject
        item.UnRead = False
        item.Save()
        return f"Marked as read: '{subject}'"

    try:
        return await bridge.call(_mark, entry_id, account)
    except Exception as e:
        return f"Error marking email as read: {format_com_error(e)}"


# =====================================================================
# TOOL 5: mark_as_unread
# =====================================================================

@mcp.tool()
async def mark_as_unread(entry_id: str, account: str = "") -> str:
    """Mark a specific email as unread in Outlook.

    Restores a previously read email to unread status. Useful for flagging
    emails that need follow-up attention. Persisted immediately.

    Args:
        entry_id: The unique Outlook EntryID of the email. Get this from
            list_emails or search_emails results.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation message with the email subject, or an error.
    """
    def _mark(outlook, namespace, entry_id, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_MAIL, "mail item"):
            return err
        subject = item.Subject
        item.UnRead = True
        item.Save()
        return f"Marked as unread: '{subject}'"

    try:
        return await bridge.call(_mark, entry_id, account)
    except Exception as e:
        return f"Error marking email as unread: {format_com_error(e)}"


# =====================================================================
# TOOL 6: move_email
# =====================================================================

@mcp.tool()
async def move_email(
    entry_id: str,
    target_folder: str = "archive",
    account: str = "",
) -> str:
    """Move an email to a different Outlook folder.

    Moves the specified email from its current location to the target folder.
    IMPORTANT: After moving, the email gets a NEW entry_id — the old one
    becomes invalid. Common use: archiving emails after processing.

    Args:
        entry_id: The unique Outlook EntryID of the email to move.
        target_folder: Destination folder name. Default is "archive". Supports
            same names as list_emails: "archive", "inbox", "sent", "deleted"/
            "trash", "drafts", "junk"/"spam", or any custom folder name.
        account: Optional. Account display name (or substring) to resolve
            the target folder in. Default: primary account.

    Returns:
        Confirmation with email subject and destination, or an error.
    """
    def _move(outlook, namespace, entry_id, target_folder, account):
        item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_MAIL, "mail item"):
            return err
        subject = item.Subject

        store = _require_store(namespace, account)
        dest = _resolve_folder(namespace, target_folder, store)
        if not dest:
            return f"Error: Target folder '{target_folder}' not found. Use list_folders to see available folders."

        item.Move(dest)
        return f"Moved '{subject}' to {target_folder}"

    try:
        return await bridge.call(_move, entry_id, target_folder, account)
    except Exception as e:
        return f"Error moving email: {format_com_error(e)}"


# =====================================================================
# TOOL 7: draft_reply
# =====================================================================

@mcp.tool()
async def draft_reply(
    entry_id: str,
    body: str,
    reply_all: bool = False,
    account: str = "",
    attachments: str = "",
    inline_replies: str = "",
) -> str:
    """Write a reply to an email as a draft. It is NOT sent.

    BEFORE CALLING THIS: call reply_context on the message, then show the
    complete draft in the chat (a short Context block about the thread unless
    it was already summarised in this conversation, then recipients, subject
    and the full text above the quoted thread) and work on it there until the
    user is happy. Then ask explicitly, as its own question,
    whether to put it in Outlook, and call this tool only after the user says
    yes. Never call it as the first step of a "write an email" request. The
    user wants to see and edit drafts in the conversation, not discover them in
    Outlook.

    Creates Outlook's normal reply (recipients filled in, quoted history and
    signature below), puts the text above it in the user's default font and
    saves it to Drafts for the user to read and press Send themselves. It stays
    in Drafts without a window unless the open_in_outlook setting is "all"; the
    open_in_outlook tool opens it if the user asks.

    POINT-BY-POINT ANSWERS GO INLINE. When the user wants to answer a list of
    questions or points "inline", "in red" or "below each point", pass
    inline_replies: each answer is written inside the quoted message, right
    under the paragraph it answers, in the inline_color setting (dark red by
    default), the way people answer a numbered list in Outlook. body is then
    short ("See below."). Preview it in the chat as pairs: the quoted line,
    then the answer marked (red). Never offer an HTML file for this instead.

    Args:
        entry_id: The unique Outlook EntryID of the email to reply to.
        body: The reply text, as plain prose; blank lines separate paragraphs.
        reply_all: If true, address everyone on the original (sender + To/CC).
            If false (default), address only the sender.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.
        attachments: Optional. Full path(s) of file(s) on this PC to attach,
            separated by semicolons.
        inline_replies: Optional. JSON array of {"after": ..., "text": ...}.
            "after" is a short phrase copied exactly from the original message
            (reply_context's message text or read_email's body; quotes and
            spacing are forgiving); the answer goes right after the paragraph
            or list item that contains it. Keep them in the order they appear:
            each is looked for after the previous one, so a phrase that repeats
            ("Confirm ok.") resolves in turn. If an anchor is not found nothing
            is saved and the result lists it. Revise with update_draft, which
            replaces all of them.

    Returns:
        JSON with status "draft_saved", sent=false, entry_id of the draft and
        the folder it is in. Tell the user the draft is waiting for them.
    """
    def _reply(outlook, namespace, entry_id, body, reply_all, account, attachments, inline_replies):
        try:
            answers = _parse_inline_replies(inline_replies)
        except ValueError as e:
            return json.dumps({"error": str(e)})
        if answers and plain_text_mode():
            return json.dumps({"error": "inline_replies needs HTML drafts; the plain_text setting is on."})
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_MAIL, "mail item"):
            return err
        reply_item = item.ReplyAll() if reply_all else item.Reply()
        files = _attach_files(reply_item, attachments)
        if answers:
            missing = _write_inline_answers(reply_item, answers)
            if missing:
                return _inline_error(missing)
        _write_draft_text(reply_item, body)
        return _save_draft(
            reply_item, "reply_all" if reply_all else "reply",
            in_reply_to=str(item.Subject or ""), to=str(_safe_get(lambda: reply_item.To) or ""),
            cc=str(_safe_get(lambda: reply_item.CC) or ""), files_attached=files,
            inline_answers=len(answers),
        )

    try:
        return await bridge.call(_reply, entry_id, body, reply_all, account, attachments, inline_replies)
    except Exception as e:
        return f"Error writing reply draft: {format_com_error(e)}"


# =====================================================================
# TOOL 7b: draft_forward
# =====================================================================

# MAPI properties that mark an attachment as inline (a picture pasted in the body).
_PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
_PR_ATTACHMENT_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"
_OL_ATTACH_OLE = 6


def _is_inline_attachment(att) -> bool:
    """True for pictures embedded in the message body and OLE objects, which are not
    files the reader would want re-attached to a forward."""
    try:
        if int(att.Type) == _OL_ATTACH_OLE:
            return True
    except Exception:
        pass
    try:
        if att.PropertyAccessor.GetProperty(_PR_ATTACHMENT_HIDDEN):
            return True
    except Exception:
        pass
    try:
        if str(att.PropertyAccessor.GetProperty(_PR_ATTACH_CONTENT_ID) or "").strip():
            return True
    except Exception:
        pass
    return False


def _attachment_key(att) -> tuple[str, int]:
    return (str(_safe_get(lambda: att.FileName) or "").lower(), int(_safe_get(lambda: att.Size, 0) or 0))


@mcp.tool()
async def draft_forward(
    entry_id: str,
    to: str,
    body: str,
    cc: str = "",
    bcc: str = "",
    account: str = "",
    forward_latest_in_thread: bool = True,
    include_thread_attachments: bool = False,
    attachments: str = "",
) -> str:
    """Forward an email as a draft. It is NOT sent.

    BEFORE CALLING THIS: call reply_context on the message so you know the
    thread, then show the complete draft in the chat first (a short Context
    block about what is being forwarded unless the thread was already
    summarised in this conversation, recipients, which message is being
    forwarded, the full note that goes above it, and which attachments will
    ride along) and work on it there until the user is happy. Then ask explicitly, as its own question, whether to put it in
    Outlook, and call this tool only after the user says yes. Never call it as
    the first step of a "forward this" request. The user wants to see and edit
    drafts in the conversation, not discover them in Outlook.

    Uses Outlook's own Forward(), so the "FW:" subject, the From/Sent/To/Subject
    header block, the quoted message and its attachments carry over exactly as
    a manual forward would. The note is placed above all of that in the user's
    default font, with the signature handled the same way as a reply. The
    result is saved to Drafts; it opens in a window only if the open_in_outlook
    setting is "all".

    Attachments come only from the message being forwarded unless
    include_thread_attachments is set.

    Args:
        entry_id: EntryID of the email to forward (from list_emails,
            search_emails, read_email or get_conversation).
        to: Recipient email addresses, separated by semicolons.
        body: The note to put above the forwarded message, as plain prose;
            blank lines separate paragraphs.
        cc: Optional. CC recipients, separated by semicolons.
        bcc: Optional. BCC recipients, separated by semicolons.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.
        forward_latest_in_thread: If true (default) and entry_id is not the
            newest message in its conversation, the newest message is forwarded
            instead so the whole chain goes. The result says which message was
            actually forwarded. Set false to forward exactly the given message.
        include_thread_attachments: If true, attachments from the earlier
            messages in the same conversation are added to the forward too,
            skipping duplicates (same filename and size) and inline pictures.
            Default false.
        attachments: Optional. Full path(s) of file(s) on this PC to attach in
            addition, separated by semicolons.

    Returns:
        JSON with status "draft_saved", sent=false, entry_id of the draft, its
        folder, opened_in_outlook, forwarded_entry_id and forwarded_subject
        (the message actually forwarded), forwarded_is_latest, and
        attachments_added (those pulled in from earlier messages). Tell the
        user where the draft is.
    """
    def _forward(outlook, namespace, entry_id, to, body, cc, bcc, account,
                 forward_latest_in_thread, include_thread_attachments, attachments):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
            store = _safe_get(lambda: item.Parent.Store) or _require_store(namespace, "")
        if err := _check_item_class(item, _OL_CLASS_MAIL, "mail item"):
            return err

        thread: list = []
        if forward_latest_in_thread or include_thread_attachments:
            thread, _source = conversation_items(store, item, limit=200)  # oldest -> newest

        source = item
        if forward_latest_in_thread and thread:
            newest = thread[-1][1]
            if str(_safe_get(lambda: newest.EntryID) or "") != str(item.EntryID):
                source = newest

        fwd = source.Forward()
        fwd.To = to
        if cc:
            fwd.CC = cc
        if bcc:
            fwd.BCC = bcc

        attachments_added: list = []
        tmpdir = None
        if include_thread_attachments and thread:
            seen = set()
            for i in range(int(_safe_get(lambda: fwd.Attachments.Count, 0) or 0)):
                seen.add(_attachment_key(fwd.Attachments.Item(i + 1)))
            tmpdir = tempfile.mkdtemp(prefix="outlook-mcp-fwd-")
            source_id = str(_safe_get(lambda: source.EntryID) or "")
            for _path, node in thread:
                if str(_safe_get(lambda: node.EntryID) or "") == source_id:
                    continue
                for i in range(int(_safe_get(lambda: node.Attachments.Count, 0) or 0)):
                    att = node.Attachments.Item(i + 1)
                    if _is_inline_attachment(att):
                        continue
                    key = _attachment_key(att)
                    if not key[0] or key in seen:
                        continue
                    target = os.path.join(tmpdir, f"{len(attachments_added) + 1:02d}_{key[0]}")
                    try:
                        att.SaveAsFile(target)
                        added = fwd.Attachments.Add(target)
                        try:
                            added.DisplayName = str(att.FileName)
                        except Exception:
                            pass
                    except Exception as e:
                        logger.warning("Could not carry attachment %s: %s", key[0], e)
                        continue
                    seen.add(key)
                    attachments_added.append({
                        "filename": str(att.FileName),
                        "size": key[1],
                        "from_subject": str(_safe_get(lambda: node.Subject) or ""),
                        "from_entry_id": str(_safe_get(lambda: node.EntryID) or ""),
                    })

        files = _attach_files(fwd, attachments)
        _write_draft_text(fwd, body)
        source_id = str(_safe_get(lambda: source.EntryID) or "")
        newest_id = str(_safe_get(lambda: thread[-1][1].EntryID) or "") if thread else source_id
        try:
            return _save_draft(
                fwd, "forward",
                forwarded_entry_id=source_id,
                forwarded_subject=str(_safe_get(lambda: source.Subject) or ""),
                forwarded_is_latest=(source_id == newest_id),
                to=str(_safe_get(lambda: fwd.To) or ""), cc=str(_safe_get(lambda: fwd.CC) or ""),
                attachments_added=attachments_added, files_attached=files,
            )
        finally:
            if tmpdir:
                shutil.rmtree(tmpdir, ignore_errors=True)

    try:
        return await bridge.call(
            _forward, entry_id, to, body, cc, bcc, account,
            forward_latest_in_thread, include_thread_attachments, attachments,
        )
    except Exception as e:
        return f"Error writing forward draft: {format_com_error(e)}"


# =====================================================================
# TOOL 8: list_folders
# =====================================================================

@mcp.tool()
async def list_folders(folder: str = "", max_depth: int = 3, account: str = "") -> str:
    """List mail folders in the user's Outlook mailbox.

    When called with no folder argument, lists top-level folders. Provide a
    folder name to drill into its subfolders — use this to browse the full
    folder tree step by step (e.g. first call with no folder to see top-level,
    then call with folder="Inbox" to see Inbox children, then
    folder="Inbox/Projects" to go deeper).

    Folder names from this output can be used directly in list_emails,
    move_email, search_emails, etc. Use slash-delimited paths for nested
    folders (e.g. "Inbox/Receipts/2026").

    Args:
        folder: Optional. Folder to list children of. Supports folder names
            ("Inbox"), slash paths ("Inbox/Receipts"), or built-in names
            ("sent", "drafts"). When empty, lists from the mailbox root.
        max_depth: How many levels deep to recurse below the starting folder.
            Default 3. Set to 1 to see only immediate children.
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of folder objects with name, full_path, item_count,
        unread_count, and subfolders (if any).
    """
    def _list(outlook, namespace, folder, max_depth, account):
        max_depth = min(max(1, max_depth), 10)
        store = _require_store(namespace, account)

        if folder:
            start = _resolve_folder(namespace, folder, store)
            if not start:
                return json.dumps({"error": f"Folder '{folder}' not found"})
            base_path = folder
        else:
            start = store.GetRootFolder()
            base_path = ""

        def walk(f, depth, path_prefix):
            current_path = f"{path_prefix}/{f.Name}" if path_prefix else f.Name
            result = {
                "name": f.Name,
                "full_path": current_path,
                "item_count": f.Items.Count,
                "unread_count": f.UnReadItemCount,
            }
            if depth < max_depth:
                children = []
                for i in range(f.Folders.Count):
                    try:
                        child = f.Folders.Item(i + 1)
                        children.append(walk(child, depth + 1, current_path))
                    except Exception:
                        continue
                if children:
                    result["subfolders"] = children
            return result

        folders = []
        for i in range(start.Folders.Count):
            try:
                child = start.Folders.Item(i + 1)
                folders.append(walk(child, 1, base_path))
            except Exception:
                continue
        return json.dumps(folders, indent=2, default=str)

    try:
        return await bridge.call(_list, folder, max_depth, account)
    except Exception as e:
        return f"Error listing folders: {format_com_error(e)}"


# =====================================================================
# TOOL 9: search_emails
# =====================================================================

_BODY_INDEX: dict = {"checked": False, "ok": False, "detail": "not probed yet"}


def _body_index_ok(namespace) -> bool:
    """Once per process: can Outlook's indexed phrase match find words inside message
    bodies on this machine? It needs Windows Search to have indexed Outlook's bodies; on
    some PCs it silently returns nothing, so the fast operator is used only after it
    proves itself. Probes with a word from a recent message in the default inbox."""
    if _BODY_INDEX["checked"]:
        return _BODY_INDEX["ok"]
    _BODY_INDEX["checked"] = True
    try:
        store = _resolve_store(namespace, "")
        items = store.GetDefaultFolder(OL_FOLDER_INBOX).Items
        items.Sort("[ReceivedTime]", True)
        probe = None
        for item in _iter_items(items, 15):  # the oldest of these: least likely to be awaiting indexing
            try:
                words = re.findall(r"[A-Za-z]{7,}", str(item.Body or "")[:3000])
            except Exception:
                continue
            if words:
                probe = (str(item.EntryID), words[len(words) // 2])
        if probe is None:
            _BODY_INDEX["detail"] = "unknown (no recent message with body text to probe with)"
            return False
        entry_id, word = probe
        hits = items.Restrict(dasl.sql(dasl.body_filter(word, indexed=True)))
        found = any(str(_safe_get(lambda h=h: h.EntryID) or "") == entry_id for h in _iter_items(hits, 300))
        _BODY_INDEX["ok"] = found
        _BODY_INDEX["detail"] = (f"indexed (probe word {word!r} found)" if found
                                 else f"not indexed (probe word {word!r} not found)")
        return found
    except Exception as e:
        _BODY_INDEX["detail"] = f"unknown (probe failed: {e})"
        return False


@mcp.tool()
async def search_emails(
    query: str = "",
    folder: str = "inbox",
    count: int = 10,
    start_date: str = "",
    end_date: str = "",
    sender: str = "",
    include_subfolders: bool = False,
    include_body: bool = False,
    match: str = "any",
    account: str = "",
    flagged_only: bool = False,
) -> str:
    """Find emails in the user's main mailbox (Outlook) by a word or name in the
    subject, sender or recipients, by sender alone, by date, and/or by
    follow-up flag. Fast: Outlook does the filtering, well under a second even
    in a 10,000-message folder.

    Start here for "find the email from X", "the thread about Y" or "my history
    with company Z". One call covers it: query= with folder="all" matches the
    subject, the sender and the To/CC/BCC of every message in the Inbox tree,
    Sent Items and Archive. A sender= filter alone lists everything from a person
    or domain. Results are newest first and each carries an entry_id for
    read_email, get_conversation, reply_context or the draft tools. Use one
    spelling and widen (folder="all", account="all", a date range) rather than
    trying variants.

    include_body=true also scans message text. That opens every message in the
    folder and takes 10 to 30 seconds per large folder; use it only when the
    normal search has failed, and add a date range. (Where Windows Search has
    indexed Outlook's bodies the server detects it and body search is fast.)

    Args:
        query: Text to find (case-insensitive substring), e.g. "Sanders",
            "EHC offer", "Earl Gray". With match="any" (default) it is looked
            for in the subject, the sender's name and address, and the To/CC/BCC
            names and addresses; with match="subject" in the subject only. With
            include_body also in the message text. May be empty when 'sender'
            is given.
        folder: Folder to search in. Default "inbox". Supports the same names
            and slash paths as list_emails, or "all" for the Inbox tree, Sent
            Items and Archive together (subfolders included).
        count: Maximum results to return. Default 10, max 200.
        start_date: Optional. Only emails received on or after this date.
            ISO 8601 (e.g. "2026-03-10" or "2026-03-10 09:00").
        end_date: Optional. Only emails received on or before this date.
            ISO 8601. Default: now (if start_date is provided).
        sender: Optional. Case-insensitive substring of the sender's name or
            email address, e.g. "castelli", "lou@pennspring.com" or "pennspring".
            Matches the real SMTP address even for senders inside the user's own
            organisation.
        include_subfolders: If true, also search every subfolder of 'folder'.
            Default false (implied by folder="all").
        include_body: If true, also match 'query' against message text. Slow on
            big folders (see above). Default false.
        match: "any" (default) or "subject". See 'query'. A short common word
            with "any" matches more, so for a word that is only ever a subject
            ("invoice") "subject" is tighter.
        account: Optional. Account display name (or substring), or "all" for
            every account (shared mailboxes not cached locally are skipped and
            listed under accounts_skipped). Default: the default account.
        flagged_only: If true, only emails flagged for follow-up in Outlook and
            not yet marked complete. May stand alone (no query or sender) for
            "my flagged emails". Default false.

    Returns:
        JSON array of matching email summaries. If the time limit cut the search
        short, accounts were skipped, or the call took more than a couple of
        seconds, an object {"results": [...], "partial": true, "note": ...,
        "accounts_skipped": [...], "elapsed_s": ...} instead.
    """
    def _search(outlook, namespace, query, folder, count, start_date, end_date,
                sender, include_subfolders, include_body, match, account, flagged_only):
        t0 = time.perf_counter()
        count = min(max(1, count), 200)
        if not query and not sender and not flagged_only:
            return json.dumps({"error": "Provide a query, a sender, flagged_only, or a combination"})
        mode = (match or "any").strip().lower()
        if mode not in ("any", "subject"):
            return json.dumps({"error": f"match must be \"any\" or \"subject\", got {match!r}"})
        stores, skipped = _stores_for(namespace, account)

        dasl_parts = []
        if query:
            with_body = include_body or _legacy()
            indexed = bool(with_body and (not _legacy()) and _body_index_ok(namespace))
            if _legacy() or mode == "subject":
                dasl_parts.append(dasl.subject_or_body_filter(query, indexed_body=indexed) if with_body
                                  else dasl.subject_filter(query))
            else:
                dasl_parts.append(dasl.anywhere_filter(query, indexed_body=indexed if with_body else None))
        if start_date:
            dasl_parts.append(dasl.received_since(_parse_date(start_date)))
        if end_date:
            dasl_parts.append(dasl.received_until(_parse_date(end_date)))
        elif start_date:
            dasl_parts.append(dasl.received_until(datetime.now()))
        if flagged_only:
            dasl_parts.append(dasl.flagged_filter())

        filter_str = dasl.sql(*dasl_parts)
        results, partial = _collect_across_stores(namespace, stores, folder, include_subfolders,
                                                  filter_str, count, sender)
        if isinstance(results, str):
            return results
        return _results_json(results, partial, skipped, time.perf_counter() - t0)

    try:
        return await bridge.call(_search, query, folder, count, start_date, end_date,
                                 sender, include_subfolders, include_body, match, account, flagged_only)
    except Exception as e:
        return f"Error searching emails: {format_com_error(e)}"


# =====================================================================
# CALENDAR TOOLS
# =====================================================================


# --- Helper: parse ISO date string ---

def _parse_date(date_str: str) -> datetime:
    """Parse ISO 8601 date string like '2026-02-25 14:00' or '2026-02-25T14:00:00'."""
    return datetime.fromisoformat(date_str)


# --- Helper: events from one or more calendar folders ---

_EVENT_SCAN_LIMIT = 2000  # occurrences inspected per calendar


def _events_in(calendars, start: datetime, end: datetime, count: int, subject_query: str = "") -> list:
    """Event summaries across ``calendars`` overlapping [start, end], sorted by start."""
    query_lower = subject_query.lower()
    results = []
    for label, folder in calendars:
        try:
            items = folder.Items
            # CRITICAL ORDER: Sort BEFORE IncludeRecurrences BEFORE Restrict
            items.Sort("[Start]")
            items.IncludeRecurrences = True
            # Overlap test, so events that started earlier but run into the window are included.
            restrict = (
                f"[Start] <= '{end.strftime('%m/%d/%Y %H:%M')}' "
                f"AND [End] >= '{start.strftime('%m/%d/%Y %H:%M')}'"
            )
            filtered = items.Restrict(restrict)
        except Exception as e:  # one unreadable calendar (e.g. an offline shared mailbox) must not sink the rest
            logger.warning("Calendar '%s' unreadable: %s", label, e)
            calendars_mod.invalidate()  # the cached folder may be stale
            results.append({"calendar": label, "error": format_com_error(e)})
            continue
        seen: set[tuple[str, str]] = set()
        for item in _iter_items(filtered, _EVENT_SCAN_LIMIT):
            if deadline.expired():
                break
            try:
                if int(item.Class) != _OL_CLASS_APPOINTMENT:
                    continue
                if query_lower and query_lower not in (item.Subject or "").lower():
                    continue
                key = (str(item.EntryID), str(item.Start))  # occurrences share the master's EntryID
                if key in seen:
                    continue
                seen.add(key)
                results.append(format_event_summary(item, label))
            except Exception:
                continue
    results.sort(key=lambda r: (r.get("start") or "", r.get("end") or ""))
    return results[:count]


def _meetings_with(calendars, start: datetime, end: datetime, needle_lower: str, limit: int = 500) -> tuple[list, bool]:
    """``(events, partial)``: meetings in [start, end] whose organizer or attendee strings
    contain ``needle_lower``. Outlook first drops plain appointments (no attendees), then
    only three fields are read per remaining item and only matches are fully formatted.
    contact_history used to format up to 2,000 occurrences per calendar before matching."""
    results: list = []
    partial = False
    for label, folder in calendars:
        if deadline.expired():
            partial = True
            break
        try:
            items = folder.Items
            items.Sort("[Start]")
            items.IncludeRecurrences = True
            restrict = (f"[Start] <= '{dasl.jet_date(end)}' AND [End] >= '{dasl.jet_date(start)}' "
                        f"AND [MeetingStatus] <> {OL_NON_MEETING}")
            filtered = items.Restrict(restrict)
        except Exception as e:
            logger.warning("Calendar '%s' unreadable: %s", label, e)
            calendars_mod.invalidate()
            continue
        seen: set[tuple[str, str]] = set()
        for item in _iter_items(filtered, _EVENT_SCAN_LIMIT):
            if deadline.expired():
                partial = True
                break
            try:
                who = " ".join(str(_safe_get(lambda a=a: getattr(item, a)) or "")
                               for a in ("Organizer", "RequiredAttendees", "OptionalAttendees")).lower()
                if needle_lower not in who:
                    continue
                key = (str(item.EntryID), str(item.Start))
                if key in seen:
                    continue
                seen.add(key)
                results.append(format_event_summary(item, label))
                if len(results) >= limit:
                    return results, partial
            except Exception:
                continue
    return results, partial


# =====================================================================
# TOOL: list_calendars
# =====================================================================

@mcp.tool()
async def list_calendars(account: str = "") -> str:
    """List every calendar folder in an account, not just the default one.

    Outlook keeps extra calendars (Birthdays, holidays, calendars shared with
    you, internet subscriptions) as separate folders, usually nested under the
    main Calendar. Each entry has path, name, is_default and item_count. Pass
    the path or name as the 'calendar' argument of list_events / search_events,
    or 'all' to read every calendar in the account at once.

    Args:
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of calendar objects, default calendar first.
    """
    def _list(outlook, namespace, account):
        store = _require_store(namespace, account)
        return json.dumps(describe_calendars(store), indent=2, default=str)

    try:
        return await bridge.call(_list, account)
    except Exception as e:
        return f"Error listing calendars: {format_com_error(e)}"


# =====================================================================
# TOOL 10: list_events
# =====================================================================

@mcp.tool()
async def list_events(
    start_date: str = "",
    end_date: str = "",
    count: int = 20,
    calendar: str = "",
    account: str = "",
) -> str:
    """List events from the user's primary calendar (Outlook) within a date range.

    Returns a JSON array of event summaries sorted by start time, including
    recurring occurrences and events that started before the range but run
    into it. Each summary has entry_id, subject, start, end (ISO 8601 with
    offset), duration, location, organizer, attendees, status info and the
    calendar it came from.

    To check availability across several calendars, call with calendar="all"
    (or once per account) and look for gaps.

    Use entry_id from results with get_event, update_event or delete_event.
    Invitations cannot be accepted or declined through this server; the
    user does that in Outlook.

    Args:
        start_date: Start of date range in ISO 8601 format (e.g. "2026-02-25"
            or "2026-02-25 09:00"). Default: now.
        end_date: End of date range. Default: 7 days from start_date.
        count: Maximum number of events to return. Default 20, max 200.
        calendar: Optional. "" (default) = the account's main Calendar; "all" =
            every calendar in the account; otherwise a calendar name or path
            from list_calendars (e.g. "Birthdays" or "Calendar/Family").
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of event summary objects.
    """
    def _list(outlook, namespace, start_date, end_date, count, calendar, account):
        count = min(max(1, count), 200)
        store = _require_store(namespace, account)
        calendars = resolve_calendars(store, calendar)
        start = _parse_date(start_date) if start_date else datetime.now()
        end = _parse_date(end_date) if end_date else start + timedelta(days=7)
        return json.dumps(_events_in(calendars, start, end, count), indent=2, default=str)

    try:
        return await bridge.call(_list, start_date, end_date, count, calendar, account)
    except Exception as e:
        return f"Error listing events: {format_com_error(e)}"


# =====================================================================
# TOOL 11: get_event
# =====================================================================

@mcp.tool()
async def get_event(entry_id: str, account: str = "") -> str:
    """Read the full details of a specific calendar event.

    Retrieves complete event information including body/description,
    attendees, recurrence status, reminders, and response status.

    Args:
        entry_id: The unique Outlook EntryID of the event. Get this from
            list_events or search_events results.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        JSON object with full event details.
    """
    def _get(outlook, namespace, entry_id, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        return json.dumps(format_event_full(item), indent=2, default=str)

    try:
        return await bridge.call(_get, entry_id, account)
    except Exception as e:
        return f"Error reading event: {format_com_error(e)}"


# =====================================================================
# TOOL 12: create_event
# =====================================================================

@mcp.tool()
async def create_event(
    subject: str,
    start: str,
    end: str,
    location: str = "",
    body: str = "",
    all_day: bool = False,
    reminder_minutes: int = 15,
    categories: str = "",
    account: str = "",
) -> str:
    """Create a personal appointment (no attendees) on the user's primary calendar (Outlook).

    Creates and saves an appointment on the user's calendar. This is a
    personal event with no attendees. Use draft_meeting instead if you need
    to invite people (the invitations are prepared, not sent).

    Args:
        subject: The event title.
        start: Start time in ISO 8601 format. Examples: "2026-02-25 14:00",
            "2026-02-25T14:00:00". For all-day events, use just the date:
            "2026-02-25".
        end: End time in ISO 8601 format. For all-day events, use the next
            day: "2026-02-26".
        location: Optional. Event location text (e.g. "Conference Room A").
            This server cannot create a Teams link; see draft_meeting.
        body: Optional. Description or notes for the event.
        all_day: If true, creates an all-day event. Default false.
        reminder_minutes: Minutes before the event to show a reminder.
            Default 15. Set to 0 to disable reminder.
        categories: Optional. Colour category name(s) from list_categories for
            this account, separated by semicolons (e.g. "Work" or
            "Work; Travel"). Each account has its own list; a name not on it
            is still applied but shows without a colour, and the result
            lists it under unknown_categories.
        account: Optional. Account display name (or substring) to create
            the event in. Default: the default account.

    Returns:
        Confirmation with event subject, entry_id and categories, or an error.
    """
    def _create(outlook, namespace, subject, start, end, location, body,
                all_day, reminder_minutes, categories, account):
        store = _require_store(namespace, account)
        appt = store.GetDefaultFolder(OL_FOLDER_CALENDAR).Items.Add(OL_APPOINTMENT_ITEM)
        appt.Subject = subject
        appt.Start = start
        appt.End = end
        if location:
            appt.Location = location
        if body:
            appt.Body = body
        appt.AllDayEvent = all_day
        if reminder_minutes > 0:
            appt.ReminderSet = True
            appt.ReminderMinutesBeforeStart = reminder_minutes
        else:
            appt.ReminderSet = False
        unknown = []
        if categories:
            appt.Categories = categories
            unknown = _unknown_categories(namespace, store, categories)
        appt.Save()
        out = {
            "status": "created",
            "subject": appt.Subject,
            "start": str(appt.Start),
            "end": str(appt.End),
            "calendar_account": str(store.DisplayName),
            "categories": str(_safe_get(lambda: appt.Categories) or ""),
            "entry_id": appt.EntryID,
        }
        if unknown:
            out["unknown_categories"] = unknown
            out["note"] = "Those categories are not on this account's list, so they show without a colour."
        return json.dumps(out, indent=2, default=str)

    try:
        return await bridge.call(
            _create, subject, start, end, location, body, all_day,
            reminder_minutes, categories, account,
        )
    except Exception as e:
        return f"Error creating event: {format_com_error(e)}"


# =====================================================================
# TOOL 13: draft_meeting
# =====================================================================

@mcp.tool()
async def draft_meeting(
    subject: str,
    start: str,
    end: str,
    required_attendees: str,
    location: str = "",
    body: str = "",
    optional_attendees: str = "",
    categories: str = "",
    account: str = "",
) -> str:
    """Prepare a meeting with attendees on the user's primary calendar (Outlook).
    Invitations are NOT sent.

    BEFORE CALLING THIS: show the complete invitation in the chat first (attendees,
    subject, time, location, text) and work on it there until the user is happy. Then ask
    explicitly, as its own question, whether to put it in Outlook, and call this
    tool only after the user says yes. Never call it as the first step of a
    "write an email" request. The user wants to see and edit invitations in the
    conversation, not discover them in Outlook.

    The meeting is saved on the user's calendar with the attendees filled in
    and invitations unsent (Outlook marks it "invitations have not been sent")
    for the user to check and press Send themselves. By default it opens in an
    Outlook window so it is not lost on the calendar (open_in_outlook setting
    "meetings"; "none" keeps it closed).

    Teams: this server cannot add a Teams meeting link; Outlook's Teams add-in
    creates those. If the user wants one, say that they click "Teams Meeting"
    in the window that opens before pressing Send, or turn on Outlook's own
    File > Options > Calendar > "Add online meeting to all meetings", which
    then does it for every meeting they send.

    Args:
        subject: The meeting title.
        start: Start time in ISO 8601 format (e.g. "2026-02-25 14:00").
        end: End time in ISO 8601 format (e.g. "2026-02-25 15:00").
        required_attendees: Required attendee email addresses, separated by
            semicolons. Example: "alice@example.com; bob@example.com"
        location: Optional. Meeting location (e.g. "Teams", "Room 301").
        body: Optional. Meeting description or agenda.
        optional_attendees: Optional. Optional attendee emails, separated
            by semicolons.
        categories: Optional. Colour category name(s) from list_categories for
            this account, separated by semicolons. A name not on the
            account's list is applied without a colour and reported under
            unknown_categories.
        account: Optional. Account display name (or substring) whose calendar
            the meeting goes on and which it will be sent from. Default: primary
            account. Use list_accounts to see available accounts.

    Returns:
        JSON with status "draft_saved", sent=false, entry_id and the calendar
        folder. Tell the user the invitation is waiting for them to send.
    """
    def _draft(outlook, namespace, subject, start, end, required_attendees,
               location, body, optional_attendees, categories, account):
        store = _require_store(namespace, account)
        appt = store.GetDefaultFolder(OL_FOLDER_CALENDAR).Items.Add(OL_APPOINTMENT_ITEM)
        if account:
            _set_sending_account(appt, outlook, store)
        appt.Subject = subject
        appt.Start = start
        appt.End = end
        appt.MeetingStatus = OL_MEETING
        if location:
            appt.Location = location
        if body:
            appt.Body = body

        for addr in required_attendees.split(";"):
            addr = addr.strip()
            if addr:
                recip = appt.Recipients.Add(addr)
                recip.Type = OL_REQUIRED

        if optional_attendees:
            for addr in optional_attendees.split(";"):
                addr = addr.strip()
                if addr:
                    recip = appt.Recipients.Add(addr)
                    recip.Type = OL_OPTIONAL

        appt.Recipients.ResolveAll()
        extra: dict = {}
        if categories:
            appt.Categories = categories
            unknown = _unknown_categories(namespace, store, categories)
            extra["categories"] = categories
            if unknown:
                extra["unknown_categories"] = unknown
        return _save_draft(
            appt, "meeting",
            start=str(appt.Start), end=str(appt.End),
            required_attendees=required_attendees, optional_attendees=optional_attendees,
            **extra,
        )

    try:
        return await bridge.call(
            _draft, subject, start, end, required_attendees, location, body,
            optional_attendees, categories, account,
        )
    except Exception as e:
        return f"Error preparing meeting: {format_com_error(e)}"


# =====================================================================
# TOOL 14: update_event
# =====================================================================

@mcp.tool()
async def update_event(
    entry_id: str,
    subject: str = "",
    start: str = "",
    end: str = "",
    location: str = "",
    body: str = "",
    categories: str = "",
    account: str = "",
) -> str:
    """Update an existing calendar event.

    Modifies properties of an appointment or meeting. Only the fields you
    provide will be updated; omitted fields remain unchanged. The change is
    saved, not sent: for a meeting the user organises, attendees are not
    notified until the user opens it in Outlook and presses Send Update.

    Args:
        entry_id: The unique Outlook EntryID of the event to update.
        subject: Optional. New event title.
        start: Optional. New start time in ISO 8601 format.
        end: Optional. New end time in ISO 8601 format.
        location: Optional. New location.
        body: Optional. New description/notes.
        categories: Optional. Colour category name(s) from list_categories
            for this account, separated by semicolons; replaces the event's
            categories. "-" clears them. A name not on the account's list is
            applied without a colour and reported under unknown_categories.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation with updated event details, or an error.
    """
    def _update(outlook, namespace, entry_id, subject, start, end, location, body, categories, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
            store = _safe_get(lambda: item.Parent.Store) or _require_store(namespace, "")
        if err := _check_item_class(item, _OL_CLASS_APPOINTMENT, "appointment/meeting item"):
            return err
        if subject:
            item.Subject = subject
        if start:
            item.Start = start
        if end:
            item.End = end
        if location:
            item.Location = location
        if body:
            item.Body = body
        unknown = []
        if categories:
            item.Categories = "" if categories.strip() == "-" else categories
            unknown = _unknown_categories(namespace, store, item.Categories or "")
        item.Save()
        out = {
            "status": "updated",
            "subject": item.Subject,
            "start": str(item.Start),
            "end": str(item.End),
            "location": item.Location or "",
            "categories": str(_safe_get(lambda: item.Categories) or ""),
            "entry_id": item.EntryID,
        }
        if unknown:
            out["unknown_categories"] = unknown
            out["note"] = "Those categories are not on this account's list, so they show without a colour."
        return json.dumps(out, indent=2, default=str)

    try:
        return await bridge.call(
            _update, entry_id, subject, start, end, location, body, categories, account,
        )
    except Exception as e:
        return f"Error updating event: {format_com_error(e)}"


# =====================================================================
# TOOL 15: delete_event
# =====================================================================

@mcp.tool()
async def delete_event(entry_id: str, account: str = "") -> str:
    """Delete a personal appointment or a received meeting from the calendar.

    Nothing is sent. Meetings the user organises (including unsent ones made
    with draft_meeting) are refused, because cancelling them would send
    notices to the attendees; the user cancels or discards those in Outlook.
    Deleting a received meeting removes it from the calendar without sending
    a response to the organiser.

    Args:
        entry_id: The unique Outlook EntryID of the event to delete.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation with the event subject, or an error.
    """
    def _delete(outlook, namespace, entry_id, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_APPOINTMENT, "appointment/meeting item"):
            return err
        subject = item.Subject

        if item.MeetingStatus == OL_MEETING:
            return (
                f"Error: '{subject}' is a meeting the user organises. Deleting it would send "
                "cancellation notices, which this server never does. Ask the user to cancel "
                "or discard it in Outlook."
            )

        item.Delete()
        return f"Event deleted: '{subject}'"

    try:
        return await bridge.call(_delete, entry_id, account)
    except Exception as e:
        return f"Error deleting event: {format_com_error(e)}"


# (There is deliberately no respond_to_meeting: accepting or declining an
#  invitation sends a message to the organiser, and this server never sends.)


# =====================================================================
# TOOL: draft_forward_event
# =====================================================================

def _add_recipients(item, addresses: str, kind: int) -> list[str]:
    """Add each semicolon-separated address to ``item.Recipients`` as ``kind``
    (required / optional attendee for an appointment)."""
    added = []
    for raw in (addresses or "").split(";"):
        addr = raw.strip()
        if addr:
            recipient = item.Recipients.Add(addr)
            recipient.Type = kind
            added.append(addr)
    return added


@mcp.tool()
async def draft_forward_event(
    entry_id: str,
    to: str,
    body: str = "",
    cc: str = "",
    account: str = "",
) -> str:
    """Forward a calendar event to someone, without sending anything.

    BEFORE CALLING THIS: show in the chat what will be forwarded (the event's
    subject, time and location), to whom, and the note that goes with it, and
    ask explicitly whether to put it in Outlook; call this tool only after the
    user says yes.

    What happens depends on the event, and the result says which:
    - A meeting the user organises: the recipients are added as attendees and
      the meeting is saved, with nothing sent. It opens in an Outlook window
      (open_in_outlook setting "meetings") where the user presses Send
      Update; Outlook offers to send it only to the added attendees. The note
      is not written into the meeting, because an update goes to everyone.
    - A meeting someone else organises, or a plain appointment: an email with
      the event attached as an iCalendar (.ics) file, saved to Drafts with the
      note above it. The recipient adds it to their calendar from the
      attachment; their accept or decline does not reach the organiser.
    Outlook's own Forward button on a received meeting (a real invitation the
    recipient can accept) exists only in Outlook's window; automation cannot
    produce it. Say so if the user wants that, and that it is one click there.

    Args:
        entry_id: EntryID of the event (from list_events, search_events or
            get_event). For a recurring series this is the whole series.
        to: Recipient email addresses, separated by semicolons.
        body: Optional note to put above the .ics email. Ignored for a
            meeting the user organises (see above).
        cc: Optional. CC recipients, separated by semicolons (optional
            attendees for a meeting the user organises).
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        JSON with status "draft_saved", sent=false, forwarded_as
        ("attendees_added" or "icalendar_attachment"), entry_id and folder of
        the saved item, and the event's subject, start, end and location.
        Tell the user where it is and what they press.
    """
    def _forward(outlook, namespace, entry_id, to, body, cc, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_APPOINTMENT, "appointment/meeting item"):
            return err
        if not (to or "").strip():
            return json.dumps({"error": "Give at least one recipient in 'to'."})
        status = int(_safe_get(lambda: item.MeetingStatus, OL_NON_MEETING) or 0)
        event = {
            "event_subject": str(_safe_get(lambda: item.Subject) or ""),
            "event_start": str(_safe_get(lambda: item.Start) or ""),
            "event_end": str(_safe_get(lambda: item.End) or ""),
            "event_location": str(_safe_get(lambda: item.Location) or ""),
        }
        if status == OL_MEETING:
            # The user organises it, so "forward" means "invite them too". Outlook's object
            # model has no Forward for appointments (only ForwardAsVcal), and adding the
            # attendee is what Outlook itself does underneath; Send Update stays with the user.
            added_to = _add_recipients(item, to, OL_REQUIRED)
            added_cc = _add_recipients(item, cc, OL_OPTIONAL)
            try:
                item.Recipients.ResolveAll()
            except Exception:
                pass
            return _save_draft(
                item, "meeting", where="on the calendar with the new attendees not yet notified",
                forwarded_as="attendees_added", to="; ".join(added_to), cc="; ".join(added_cc),
                note_applied=False,
                next_step="The user opens the meeting and presses Send Update; Outlook offers to send it "
                          "only to the added attendees.",
                **event,
            )
        mail = item.ForwardAsVcal()  # a MailItem with the event attached as .ics
        mail.To = to
        if cc:
            mail.CC = cc
        if body:
            _write_draft_text(mail, body)
        return _save_draft(mail, "forward_event", forwarded_as="icalendar_attachment",
                           to=to, cc=cc, note_applied=bool(body), **event)

    try:
        return await bridge.call(_forward, entry_id, to, body, cc, account)
    except Exception as e:
        return f"Error forwarding event: {format_com_error(e)}"


# =====================================================================
# TOOL: move_event
# =====================================================================

@mcp.tool()
async def move_event(
    entry_id: str,
    calendar: str,
    account: str = "",
    to_account: str = "",
) -> str:
    """Move a calendar item to another calendar: a secondary calendar in the
    same mailbox (for example "Trips"), or a calendar in another account.
    Nothing is sent.

    For "put the whole trip on my other calendar": list the events first
    (list_events or search_events over the trip's dates), show the user the
    list and confirm it, then call this once per event. Each call returns the
    item's new entry_id; the old one stops working.

    Two things the user should know; say them when they apply:
    - Classic Outlook fires reminders and reports free/busy only from the
      main Calendar of a mailbox. Items on a secondary calendar show there but
      do not pop reminders.
    - A meeting moved to a different account becomes a plain copy there.
      Updates or cancellations from its organiser keep going to the original
      mailbox and do not follow the copy. Fine for flights and holds, not for
      meetings someone else organises. Moving within one mailbox keeps
      everything intact.

    Args:
        entry_id: EntryID of the event (from list_events, search_events or
            get_event). For a recurring series this moves the whole series.
        calendar: Destination calendar: a name or path from list_calendars
            ("Trips", "Calendar/Trips"), or "" for the main Calendar.
        account: Optional. Account the event is in now (display name or
            substring). Only needed if entry_id is ambiguous across stores.
        to_account: Optional. Account that holds the destination calendar.
            Default: the account the event is in.

    Returns:
        JSON with status "moved" (or "unchanged" if it was already there),
        the subject, from_calendar, to_calendar, to_account, the new entry_id
        and a note with the caveats that apply.
    """
    def _move(outlook, namespace, entry_id, calendar, account, to_account):
        if (calendar or "").strip().lower() in ("all", "*"):
            return json.dumps({"error": "calendar must name one destination calendar, not \"all\"."})
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
            store = _safe_get(lambda: item.Parent.Store) or _require_store(namespace, "")
        if err := _check_item_class(item, _OL_CLASS_APPOINTMENT, "appointment/meeting item"):
            return err
        dest_store = _require_store(namespace, to_account) if to_account else store
        label, dest = resolve_calendars(dest_store, calendar)[0]
        subject = str(_safe_get(lambda: item.Subject) or "")
        from_path = folder_path(item)
        if str(_safe_get(lambda: item.Parent.EntryID) or "") == str(dest.EntryID):
            return json.dumps({"status": "unchanged", "subject": subject, "entry_id": str(item.EntryID),
                               "note": f"Already on {label}."}, indent=2, default=str)
        cross_account = str(dest_store.StoreID) != str(store.StoreID)
        is_meeting = int(_safe_get(lambda: item.MeetingStatus, OL_NON_MEETING) or 0) != OL_NON_MEETING
        try:
            dest_is_main = str(dest.EntryID) == str(dest_store.GetDefaultFolder(OL_FOLDER_CALENDAR).EntryID)
        except Exception:
            dest_is_main = False
        moved = item.Move(dest)
        note = "Nothing was sent."
        if not dest_is_main:
            note += " Items on a secondary calendar do not pop reminders in classic Outlook."
        if cross_account and is_meeting:
            note += (" This meeting is now a plain copy in the other account; updates or cancellations "
                     "from its organiser will not reach it.")
        return json.dumps({
            "status": "moved",
            "subject": subject,
            "from_calendar": from_path,
            "to_calendar": label,
            "to_account": str(dest_store.DisplayName),
            "entry_id": str(_safe_get(lambda: moved.EntryID) or ""),
            "note": note,
        }, indent=2, default=str)

    try:
        return await bridge.call(_move, entry_id, calendar, account, to_account)
    except Exception as e:
        return f"Error moving event: {format_com_error(e)}"


# =====================================================================
# TOOL 16: search_events
# =====================================================================

@mcp.tool()
async def search_events(
    query: str,
    start_date: str = "",
    end_date: str = "",
    count: int = 10,
    calendar: str = "",
    account: str = "",
) -> str:
    """Search the user's primary calendar (Outlook) for events by keyword.

    Searches event subjects within a date range. Results are sorted by
    start time. Includes recurring event occurrences.

    Args:
        query: The search term (case-insensitive substring match on subject).
            Examples: "standup", "review", "1:1".
        start_date: Start of search range in ISO 8601 format. Default: 30
            days ago.
        end_date: End of search range. Default: 30 days from now.
        count: Maximum results to return. Default 10, max 200.
        calendar: Optional. "" (default) = the account's main Calendar; "all" =
            every calendar in the account; otherwise a name or path from
            list_calendars.
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of matching event summaries; when the call took more than a
        couple of seconds (a long date range over many calendars expands every
        recurrence), an object {"results": [...], "elapsed_s": ...} instead.
    """
    def _search(outlook, namespace, query, start_date, end_date, count, calendar, account):
        t0 = time.perf_counter()
        count = min(max(1, count), 200)
        store = _require_store(namespace, account)
        calendars = resolve_calendars(store, calendar)
        start = _parse_date(start_date) if start_date else datetime.now() - timedelta(days=30)
        end = _parse_date(end_date) if end_date else datetime.now() + timedelta(days=30)
        events = _events_in(calendars, start, end, count, query)
        return _results_json(events, False, [], time.perf_counter() - t0)

    try:
        return await bridge.call(_search, query, start_date, end_date, count, calendar, account)
    except Exception as e:
        return f"Error searching events: {format_com_error(e)}"


# =====================================================================
# TASK TOOLS
# =====================================================================

@mcp.tool()
async def list_tasks(
    include_completed: bool = False,
    count: int = 20,
    account: str = "",
) -> str:
    """List tasks from the user's primary task list (the Outlook Tasks folder).

    Returns a JSON array of task summaries sorted by due date. Each task
    includes entry_id, subject, status, percent_complete, due_date,
    importance, and categories.

    Args:
        include_completed: If true, include completed tasks. Default false
            (only pending/in-progress tasks).
        count: Maximum number of tasks to return. Default 20.
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of task summary objects.
    """
    def _list(outlook, namespace, include_completed, count, account):
        count = min(max(1, count), 200)
        store = _require_store(namespace, account)
        folder = store.GetDefaultFolder(OL_FOLDER_TASKS)
        items = folder.Items
        items.Sort("[DueDate]")

        if not include_completed:
            items = items.Restrict("[Complete] = False")

        results = []
        limit = min(count, items.Count)
        for i in range(limit):
            try:
                results.append(format_task_summary(items.Item(i + 1)))
            except Exception:
                continue
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, include_completed, count, account)
    except Exception as e:
        return f"Error listing tasks: {format_com_error(e)}"


@mcp.tool()
async def get_task(entry_id: str, account: str = "") -> str:
    """Read the full details of a specific task.

    Args:
        entry_id: The unique Outlook EntryID of the task.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        JSON object with full task details including body.
    """
    def _get(outlook, namespace, entry_id, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        return json.dumps(format_task_full(item), indent=2, default=str)

    try:
        return await bridge.call(_get, entry_id, account)
    except Exception as e:
        return f"Error reading task: {format_com_error(e)}"


@mcp.tool()
async def create_task(
    subject: str,
    body: str = "",
    due_date: str = "",
    importance: str = "normal",
    reminder_minutes: int = 0,
    account: str = "",
) -> str:
    """Create a new task in the user's primary task list (Outlook).

    Args:
        subject: The task title.
        body: Optional. Task description or notes.
        due_date: Optional. Due date in ISO 8601 format (e.g. "2026-03-01").
        importance: Optional. "low", "normal" (default), or "high".
        reminder_minutes: Optional. Minutes before due date to remind.
            Default 0 (no reminder).
        account: Optional. Account display name (or substring) to create
            the task in. Default: primary account.

    Returns:
        Confirmation with task subject and entry_id.
    """
    def _create(outlook, namespace, subject, body, due_date, importance,
                reminder_minutes, account):
        task = outlook.CreateItem(OL_TASK_ITEM)
        # Move to correct store's tasks folder if account specified
        if account:
            store = _require_store(namespace, account)
            tasks_folder = store.GetDefaultFolder(OL_FOLDER_TASKS)
            task.Move(tasks_folder)
            task = namespace.GetItemFromID(task.EntryID)
        task.Subject = subject
        if body:
            task.Body = body
        if due_date:
            task.DueDate = due_date
        imp_map = {"low": 0, "normal": 1, "high": 2}
        task.Importance = imp_map.get(importance.lower(), 1)
        if reminder_minutes > 0:
            task.ReminderSet = True
            task.ReminderMinutesBeforeStart = reminder_minutes
        else:
            task.ReminderSet = False
        task.Save()
        return json.dumps({
            "status": "created",
            "subject": task.Subject,
            "entry_id": task.EntryID,
            "due_date": str(task.DueDate) if due_date else None,
        }, indent=2, default=str)

    try:
        return await bridge.call(
            _create, subject, body, due_date, importance, reminder_minutes,
            account,
        )
    except Exception as e:
        return f"Error creating task: {format_com_error(e)}"


@mcp.tool()
async def complete_task(entry_id: str, account: str = "") -> str:
    """Mark a task as complete.

    Sets the task status to complete and percent to 100%.

    Args:
        entry_id: The unique Outlook EntryID of the task.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation with the task subject.
    """
    def _complete(outlook, namespace, entry_id, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_TASK, "task item"):
            return err
        item.Status = OL_TASK_COMPLETE
        item.PercentComplete = 100
        item.Save()
        return f"Task completed: '{item.Subject}'"

    try:
        return await bridge.call(_complete, entry_id, account)
    except Exception as e:
        return f"Error completing task: {format_com_error(e)}"


@mcp.tool()
async def delete_task(entry_id: str, account: str = "") -> str:
    """Delete a task from Outlook.

    Args:
        entry_id: The unique Outlook EntryID of the task to delete.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation with the task subject.
    """
    def _delete(outlook, namespace, entry_id, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_TASK, "task item"):
            return err
        subject = item.Subject
        item.Delete()
        return f"Task deleted: '{subject}'"

    try:
        return await bridge.call(_delete, entry_id, account)
    except Exception as e:
        return f"Error deleting task: {format_com_error(e)}"


# =====================================================================
# ATTACHMENT TOOLS
# =====================================================================

@mcp.tool()
async def list_attachments(entry_id: str, account: str = "") -> str:
    """List all attachments on an email or calendar event.

    Args:
        entry_id: The EntryID of the email or event to check for attachments.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        JSON array of attachment objects with index, filename, and size.
    """
    def _list(outlook, namespace, entry_id, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        results = []
        for i in range(item.Attachments.Count):
            att = item.Attachments.Item(i + 1)
            results.append({
                "index": i + 1,
                "filename": att.FileName,
                "size": att.Size,
            })
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, entry_id, account)
    except Exception as e:
        return f"Error listing attachments: {format_com_error(e)}"


@mcp.tool()
async def save_attachment(
    entry_id: str,
    attachment_index: int = 1,
    save_directory: str = "",
    account: str = "",
) -> str:
    """Save an attachment from an email or event to disk.

    Downloads the specified attachment to a local directory.

    Args:
        entry_id: The EntryID of the email or event containing the attachment.
        attachment_index: Which attachment to save (1-based index). Default 1
            (first attachment). Use list_attachments to see available indices.
        save_directory: Directory to save the file to. Default: user's
            Downloads folder.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        The full file path where the attachment was saved, or an error.
    """
    def _save(outlook, namespace, entry_id, attachment_index, save_directory, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if attachment_index < 1 or item.Attachments.Count < attachment_index:
            return f"Error: Only {item.Attachments.Count} attachment(s), requested index {attachment_index}"

        att = item.Attachments.Item(attachment_index)
        if not save_directory:
            save_directory = os.path.join(os.path.expanduser("~"), "Downloads")

        # Resolve to real path before creating
        save_directory = os.path.realpath(save_directory)
        os.makedirs(save_directory, exist_ok=True)

        # Strip path separators and dangerous characters from filename
        safe_name = os.path.basename(att.FileName)
        safe_name = re.sub(r'[^\w\.\-_ ]', '_', safe_name)
        if not safe_name:
            safe_name = "attachment"

        save_path = os.path.join(save_directory, safe_name)

        # Ensure final path is still inside the intended directory
        if not os.path.realpath(save_path).startswith(save_directory + os.sep) and \
           os.path.realpath(save_path) != save_directory:
            return "Error: Attachment filename would escape the target directory."

        att.SaveAsFile(save_path)
        return json.dumps({
            "status": "saved",
            "filename": safe_name,
            "path": save_path,
            "size": att.Size,
        }, indent=2, default=str)

    try:
        return await bridge.call(_save, entry_id, attachment_index, save_directory, account)
    except Exception as e:
        return f"Error saving attachment: {format_com_error(e)}"


# =====================================================================
# CATEGORY TOOLS
# =====================================================================

_CATEGORY_COLORS = {
    0: "none", 1: "red", 2: "orange", 3: "peach", 4: "yellow", 5: "green", 6: "teal", 7: "olive",
    8: "blue", 9: "purple", 10: "maroon", 11: "steel", 12: "dark steel", 13: "gray", 14: "dark gray",
    15: "black", 16: "dark red", 17: "dark orange", 18: "dark peach", 19: "dark yellow", 20: "dark green",
    21: "dark teal", 22: "dark olive", 23: "dark blue", 24: "dark purple", 25: "dark maroon",
}


def _store_categories(namespace, store) -> list[dict]:
    """The colour categories of one mailbox. Each Exchange account keeps its own master list
    (``Store.Categories``); the profile-wide list is the fallback for stores without one."""
    cats = _safe_get(lambda: store.Categories) if store is not None else None
    if cats is None:
        cats = namespace.Categories
    out = []
    for i in range(int(_safe_get(lambda: cats.Count, 0) or 0)):
        try:
            cat = cats.Item(i + 1)
            color = int(_safe_get(lambda: cat.Color, 0) or 0)
            out.append({"name": str(cat.Name), "color": color, "color_name": _CATEGORY_COLORS.get(color, str(color))})
        except Exception:
            continue
    return out


def _split_categories(text: str) -> list[str]:
    return [c.strip() for c in re.split(r"[;,]", text or "") if c.strip()]


def _unknown_categories(namespace, store, text: str) -> list[str]:
    """Names in ``text`` that are not on the store's category list (they still apply, uncoloured)."""
    known = {c["name"].lower() for c in _store_categories(namespace, store)}
    return [c for c in _split_categories(text) if c.lower() not in known]


@mcp.tool()
async def list_categories(account: str = "") -> str:
    """List the colour categories of one account.

    Each account (mailbox) keeps its own list, so "Work" may exist in one and
    not another; ask for the account the item lives in. The names are what
    set_category, create_event, draft_meeting and update_event accept.

    Args:
        account: Optional. Account display name (or substring). Default: the
            default account. Use list_accounts to see available accounts.

    Returns:
        JSON with the account and its categories [{name, color, color_name}].
    """
    def _list(outlook, namespace, account):
        store = _require_store(namespace, account)
        return json.dumps({
            "account": str(store.DisplayName),
            "categories": _store_categories(namespace, store),
            "note": "Categories are per account; other accounts may have different ones.",
        }, indent=2, default=str)

    try:
        return await bridge.call(_list, account)
    except Exception as e:
        return f"Error listing categories: {format_com_error(e)}"


@mcp.tool()
async def set_category(
    entry_id: str,
    categories: str,
    account: str = "",
) -> str:
    """Set the colour categories on an email, event or task.

    Replaces any existing categories on the item. Use the names from
    list_categories for the account the item lives in; each account has its
    own list. A name not on that list is still applied but shows without a
    colour, and the result lists it under unknown_categories.

    Args:
        entry_id: The EntryID of the item to categorize.
        categories: Category name(s), separated by semicolons or commas.
            Example: "Work" or "Work; Follow-up". Use an empty string to
            clear all categories.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        JSON with the item's subject, the categories now on it, and
        unknown_categories when some were not on the account's list.
    """
    def _set(outlook, namespace, entry_id, categories, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
            store = _safe_get(lambda: item.Parent.Store) or _require_store(namespace, "")
        item.Categories = categories
        item.Save()
        now = str(_safe_get(lambda: item.Categories) or "")
        out = {"status": "categories_set", "subject": str(_safe_get(lambda: item.Subject) or ""),
               "categories": now, "account": str(store.DisplayName)}
        unknown = _unknown_categories(namespace, store, now)
        if unknown:
            out["unknown_categories"] = unknown
            out["note"] = "Those categories are not on this account's list, so they show without a colour."
        return json.dumps(out, indent=2, default=str)

    try:
        return await bridge.call(_set, entry_id, categories, account)
    except Exception as e:
        return f"Error setting categories: {format_com_error(e)}"


# =====================================================================
# RULES TOOLS
# =====================================================================

@mcp.tool()
async def list_rules(account: str = "") -> str:
    """List all mail rules in Outlook.

    Returns the configured inbox rules with their names and enabled status.

    Args:
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of rule objects with name, enabled status, and index.
    """
    def _list(outlook, namespace, account):
        store = _require_store(namespace, account)
        rules = store.GetRules()
        results = []
        for i in range(rules.Count):
            rule = rules.Item(i + 1)
            results.append({
                "index": i + 1,
                "name": rule.Name,
                "enabled": bool(rule.Enabled),
            })
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list, account)
    except Exception as e:
        return f"Error listing rules: {format_com_error(e)}"


@mcp.tool()
async def toggle_rule(
    rule_name: str,
    enabled: bool,
    account: str = "",
) -> str:
    """Enable or disable a mail rule by name.

    CAUTION: This modifies live mail rules immediately. Confirm the rule name
    with list_rules before calling.

    Args:
        rule_name: The exact name of the rule to toggle. Use list_rules
            to see available rule names.
        enabled: True to enable the rule, False to disable it.
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        Confirmation with the rule name and new status.
    """
    def _toggle(outlook, namespace, rule_name, enabled, account):
        store = _require_store(namespace, account)
        rules = store.GetRules()
        for i in range(rules.Count):
            rule = rules.Item(i + 1)
            if rule.Name == rule_name:
                logger.warning(
                    "toggle_rule: setting rule '%s' enabled=%s", rule_name, enabled
                )
                rule.Enabled = enabled
                rules.Save()
                status = "enabled" if enabled else "disabled"
                return f"Rule '{rule_name}' {status}"
        return f"Error: Rule '{rule_name}' not found. Use list_rules to see available rules."

    try:
        return await bridge.call(_toggle, rule_name, enabled, account)
    except Exception as e:
        return f"Error toggling rule: {format_com_error(e)}"


# =====================================================================
# OUT OF OFFICE TOOLS
# =====================================================================

@mcp.tool()
async def get_out_of_office(account: str = "") -> str:
    """Check the current Out of Office (auto-reply) status.

    Returns whether Out of Office is currently enabled.

    Args:
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON object with the OOF status.
    """
    def _get(outlook, namespace, account):
        store = _require_store(namespace, account)
        try:
            prop_tag = "http://schemas.microsoft.com/mapi/proptag/0x661D000B"
            oof_state = store.PropertyAccessor.GetProperty(prop_tag)
            return json.dumps({
                "out_of_office": bool(oof_state),
                "status": "on" if oof_state else "off",
            }, indent=2)
        except Exception:
            return json.dumps({
                "out_of_office": None,
                "status": "unknown",
                "note": "Could not read OOF property. Check Outlook settings directly.",
            }, indent=2)

    try:
        return await bridge.call(_get, account)
    except Exception as e:
        return f"Error checking OOF status: {format_com_error(e)}"


# =====================================================================
# Entry point
# =====================================================================

# =====================================================================
# SETTINGS: get_settings / set_setting
# =====================================================================

@mcp.tool()
async def get_settings() -> str:
    """Show the server's settings, where each value comes from, and which account
    tools use when none is named.

    Settings: default_account, hide_accounts, open_in_outlook, font_family,
    font_size_pt, font_color, plain_text, include_shared_in_all, call_timeout_s,
    legacy_mode. Each value comes from an environment
    variable (set in Claude's config), the settings file (written by
    set_setting), or the built-in default; "source" says which. Change one with
    set_setting.

    Returns:
        JSON with settings_file, settings {key: {value, source, env_var, help}}
        and effective_default_account.
    """
    def _get(outlook, namespace):
        store = _resolve_store(namespace, "")
        return json.dumps({
            "settings_file": str(settings.settings_path()),
            "settings": settings.all_settings(),
            "effective_default_account": str(store.DisplayName) if store is not None else None,
            "note": "A value whose source is \"env\" is set in Claude's config and overrides the settings file; "
                    "set_setting cannot change it until that env entry is removed.",
        }, indent=2, default=str)

    try:
        return await bridge.call(_get)
    except Exception as e:
        return f"Error reading settings: {format_com_error(e)}"


@mcp.tool()
async def set_setting(key: str, value: str = "", clear: bool = False) -> str:
    """Change one server setting. Takes effect immediately and survives restarts.

    Keys:
        default_account: account used when a tool is not given one, e.g.
            "make HGP my default" -> set_setting("default_account", "heartland").
            Checked against the accounts in Outlook; a unique substring is enough
            and the full display name is stored.
        hide_accounts: accounts to leave out of list_accounts and account="all",
            as a semicolon-separated list of display-name substrings.
        open_in_outlook: which saved drafts open in an Outlook window:
            "meetings" (default: unsent meeting invitations only), "all" or
            "none". Email drafts always land in Drafts; the open_in_outlook
            tool opens any item on request. (open_drafts is the older
            true/false name for the same thing.)
        font_family, font_size_pt, font_color, plain_text: how draft text is written.
        include_shared_in_all: "true" to include shared, non-cached mailboxes in
            account="all" searches (slow); default false.
        call_timeout_s: seconds a tool call may take (default 60); long searches
            return partial results just before it.
        legacy_mode: "true" to run the pre-0.6 (slow) search paths; for
            troubleshooting only.

    Args:
        key: One of the keys above.
        value: The new value. For hide_accounts, substrings separated by ";".
        clear: If true, remove the key from the settings file so the default
            (or an environment variable, if set) applies again.

    Returns:
        JSON with the key, what was saved, the value now in effect and its
        source. If an environment variable overrides this key, says so.
    """
    def _set(outlook, namespace, key, value, clear):
        key = (key or "").strip()
        if key not in settings.SETTINGS:
            return json.dumps({"error": f"Unknown setting {key!r}", "known": list(settings.SETTINGS)})
        if clear:
            settings.clear_value(key)
        else:
            value = (value or "").strip()
            if key == "default_account" and value:
                names = [str(namespace.Stores.Item(i + 1).DisplayName) for i in range(namespace.Stores.Count)]
                exact = [n for n in names if n.lower() == value.lower()]
                partial = [n for n in names if value.lower() in n.lower()]
                hits = exact or partial
                if not hits:
                    return json.dumps({"error": f"No account matches {value!r}", "accounts": names})
                if len(hits) > 1:
                    return json.dumps({"error": f"{value!r} matches more than one account; be more specific", "matches": hits})
                value = hits[0]
            if key == "hide_accounts" and value:
                names = [str(namespace.Stores.Item(i + 1).DisplayName).lower() for i in range(namespace.Stores.Count)]
                unknown = [part.strip() for part in value.split(";") if part.strip()
                           and not any(part.strip().lower() in n for n in names)]
                if unknown:
                    logger.warning("hide_accounts entries match no account right now: %s", unknown)
            try:
                settings.set_value(key, value)
            except ValueError as e:
                return json.dumps({"error": str(e), "choices": list(settings.SETTINGS[key].get("choices", []))})
        now, source = settings.get_with_source(key)
        out = {"key": key, "cleared": clear, "saved": None if clear else value, "effective": now, "source": source,
               "settings_file": str(settings.settings_path())}
        if source == "env":
            out["warning"] = (f"{settings.SETTINGS[key]['env']} is set in Claude's config and overrides this. "
                              "Remove it there (and restart Claude) for the saved value to apply.")
        return json.dumps(out, indent=2, default=str)

    try:
        return await bridge.call(_set, key, value, clear)
    except Exception as e:
        return f"Error changing setting: {format_com_error(e)}"


# =====================================================================
# TOOL: read_attachment
# =====================================================================

_OL_ATTACH_EMBEDDED_ITEM = 5


@mcp.tool()
async def read_attachment(
    entry_id: str,
    attachment_index: int = 1,
    max_chars: int = 20000,
    pages: str = "",
    sheet: str = "",
    account: str = "",
) -> str:
    """Read the text of an attachment without the user opening or saving it.

    Works for PDF, Word (.docx), Excel (.xlsx/.xlsm), PowerPoint (.pptx), CSV,
    plain text, HTML and attached Outlook messages (.msg). Old binary formats
    (.doc/.xls/.ppt) and pictures are not readable; scanned PDFs have no text
    and are reported as such. The file is copied to a temp folder for reading
    and deleted afterwards; the mailbox is not changed.

    Args:
        entry_id: EntryID of the email or event holding the attachment.
        attachment_index: Which attachment (1-based; see list_attachments).
        max_chars: Cap on returned text. Default 20000. The result says if the
            cap was hit; narrow with pages/sheet rather than raising it a lot.
        pages: For PDF and PowerPoint: which pages/slides, e.g. "1-3,7".
            Default all.
        sheet: For Excel: a sheet name (or part of one). Default all sheets,
            each capped at 500 rows.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        JSON with filename, size, kind, the extracted text, and details such as
        pages_total / sheets / truncated.
    """
    def _read(outlook, namespace, entry_id, attachment_index, max_chars, pages, sheet, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        count = int(_safe_get(lambda: item.Attachments.Count, 0) or 0)
        if count == 0:
            return json.dumps({"error": "This item has no attachments"})
        if attachment_index < 1 or attachment_index > count:
            return json.dumps({"error": f"attachment_index must be between 1 and {count}"})
        att = item.Attachments.Item(attachment_index)
        filename = str(_safe_get(lambda: att.FileName) or f"attachment{attachment_index}")
        max_chars = max(500, min(int(max_chars or 20000), 200000))
        tmpdir = tempfile.mkdtemp(prefix="outlook-mcp-read-")
        try:
            safe_name = re.sub(r"[\\/:*?\"<>|]", "_", filename) or "attachment"
            path = os.path.join(tmpdir, safe_name)
            att.SaveAsFile(path)
            result = {"filename": filename, "size": int(_safe_get(lambda: att.Size, 0) or 0), "index": attachment_index}
            if int(_safe_get(lambda: att.Type, 0) or 0) == _OL_ATTACH_EMBEDDED_ITEM or safe_name.lower().endswith(".msg"):
                # An attached Outlook message: open it as an item and read it like an email.
                try:
                    inner = namespace.OpenSharedItem(path)
                    record = format_email_full(inner, body_max_length=max_chars)
                    record["kind"] = "msg"
                    record["attachments"] = [
                        {"index": i + 1, "filename": str(inner.Attachments.Item(i + 1).FileName)}
                        for i in range(int(_safe_get(lambda: inner.Attachments.Count, 0) or 0))
                    ]
                    try:
                        inner.Close(1)  # olDiscard
                    except Exception:
                        pass
                    result.update(record)
                    return json.dumps(result, indent=2, default=str)
                except Exception as e:
                    result["error"] = f"Could not open the attached message: {format_com_error(e)}"
                    return json.dumps(result, indent=2, default=str)
            text, info = extract(path, max_chars=max_chars, pages=pages, sheet=sheet)
            result.update(info)
            result["text"] = text
            return json.dumps(result, indent=2, default=str)
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)

    try:
        return await bridge.call(_read, entry_id, attachment_index, max_chars, pages, sheet, account)
    except Exception as e:
        return f"Error reading attachment: {format_com_error(e)}"


# =====================================================================
# TOOL: contact_history
# =====================================================================

@mcp.tool()
async def contact_history(
    person: str,
    days: int = 365,
    count: int = 10,
    account: str = "all",
    include_meetings: bool = True,
) -> str:
    """Everything recent with one person in the user's Outlook mail and calendars,
    in one call: when they last wrote, when the user last wrote to them, mail
    where someone else wrote to them and the user was copied, the recent
    messages in every direction, and the last and next meetings with them.
    Looks across every account (unless one is named), the whole Inbox tree, Sent
    Items and every calendar.

    This is the broadest search the server offers. Use it for one specific
    person once you know who they are, not as the first move for "find the
    email about X" and never for a company, deal or project name: those are
    subject searches (search_emails with query= and folder="all"). If nothing
    matches a person, the result carries subject_matches instead, so a company
    name still finds its thread, but only after the full sweep came back empty.
    Give it something specific: a surname, a full name, an address or a domain.
    A short first name like "dan" matches every Dan, Dante and Jordan.

    Args:
        person: Part of the person's name or email address, e.g. "castelli",
            "lou@pennspring.com" or "pennspring.com". Case-insensitive.
        days: How far back to look. Default 365.
        count: Messages to return per direction per account (and in the merged
            timeline). Default 10, max 50.
        account: "all" (default) for every visible account (shared mailboxes not
            cached locally are skipped and listed under accounts_skipped), or
            one account's display name (or substring).
        include_meetings: Also scan calendars for meetings with this person.
            Default true.

    Returns:
        JSON with last_received, last_sent, received_count, sent_count,
        copied_count, recent (merged timeline, newest first, each tagged with a
        direction of received, sent or copied, and its account), meetings {last,
        next, past_count, upcoming_count}, accounts_searched, accounts_skipped
        and elapsed_s. subject_matches (with a note) appears only when no mail
        from or to the person was found. sent_scan_capped, with
        sent_scan_partial listing the accounts and how much of each was
        checked, appears only when the slow fallback scan of Sent Items
        stopped with messages still unread in the window. "partial": true means
        the time limit stopped the sweep; narrow it (one account, fewer days)
        rather than repeating it.
    """
    def _history(outlook, namespace, person, days, count, account, include_meetings):
        t0 = time.perf_counter()
        needle = (person or "").strip()
        if not needle:
            return json.dumps({"error": "person is required"})
        needle_lower = needle.lower()
        count = min(max(1, int(count)), 50)
        days = min(max(1, int(days)), 3650)
        now = datetime.now()
        since = now - timedelta(days=days)
        legacy = _legacy()
        if (account or "").strip().lower() == "all":
            stores, skipped = _sweep_stores(namespace)
        else:
            stores, skipped = [_require_store(namespace, account)], []
        stamp = dasl.jet_date(since)
        received: list = []
        sent: list = []
        copied: list = []
        searched = []
        partial = False
        sent_scan_partial: list = []  # per account where the Sent Items walk stopped before the end of the window
        scan_spent = 0.0
        inbox_trees: dict[str, list] = {}
        for store in stores:
            if deadline.expired():
                partial = True
                break
            name = str(store.DisplayName)
            searched.append(name)
            own = _own_addresses(outlook, store)
            # Inbox tree, matched by Outlook. The sender's name, directory address and
            # SMTP address are matched, and (unless legacy) the To/CC/BCC display strings
            # too, so mail where the person is a co-recipient (a colleague wrote to them
            # and copied the user) and the user's own replies filed outside Sent Items
            # are found as well. The sender then says which way each message went.
            folders: list = []
            try:
                _folder_subtree(store.GetDefaultFolder(OL_FOLDER_INBOX), folders)
            except Exception:
                folders = []
            inbox_trees[name] = folders
            people = dasl.sender_filter(needle) if legacy else dasl.people_filter(needle)
            found, cut = _collect_summaries(folders, dasl.sql(people, dasl.received_since(since)),
                                            count if legacy else count * 2, "")
            partial = partial or cut
            for r in found:
                r["account"] = name
                r["direction"] = "received" if legacy else classify_direction(
                    r.get("sender", ""), r.get("sender_name", ""), needle, own)
                {"received": received, "sent": sent}.get(r["direction"], copied).append(r)

            # Sent Items: Outlook matches the To/CC/BCC display strings. Those hold names,
            # so the names this person has written under are searched alongside the
            # needle (an address needle alone would miss "Lou Castelli"). Python walks
            # Sent Items only when Outlook rejects the filter, in legacy mode, or when an
            # address-like needle found nothing natively. That walk resolves every
            # recipient's address, so it is capped by count and by seconds: uncapped it
            # once took 30 s across three accounts to confirm there was nothing.
            known_names: list[str] = []
            for r in received:
                n = str(r.get("sender_name") or "").strip()
                if n and n.lower() not in (k.lower() for k in known_names) and len(known_names) < 6:
                    known_names.append(n)
            try:
                items = store.GetDefaultFolder(OL_FOLDER_SENT_MAIL).Items
                items.Sort("[SentOn]", True)
                items = items.Restrict(f"[SentOn] >= '{stamp}'")
            except Exception:
                items = None
            sent_here: list = []
            native_ok = False
            if items is not None and not legacy:
                try:
                    native = items.Restrict(dasl.sql(dasl.recipient_filter([needle] + known_names)))
                    native_ok = True
                    for item in _iter_items(native, count):
                        if deadline.expired():
                            partial = True
                            break
                        try:
                            if int(item.Class) != _OL_CLASS_MAIL:
                                continue
                            sent_here.append(_sent_summary(item, name))
                        except Exception:
                            continue
                except Exception as e:
                    logger.warning("Native recipient filter rejected (%s); scanning Sent Items in Python", e)
                    native_ok = False
            # The walk is pointless when the Inbox tree already showed this address raw in
            # a To/CC string (Outlook would then have matched Sent Items natively too), and
            # it is budgeted per account and per call. Legacy mode keeps the old, uncapped walk.
            scan_python = items is not None and (
                legacy or not native_ok
                or (not sent_here and dasl.looks_like_address(needle) and not found))
            if scan_python and not legacy and scan_spent >= _SENT_SCAN_TOTAL_S:
                scan_python = False
                sent_scan_partial.append({"account": name, "scanned": 0,
                                          "of": int(_safe_get(lambda: items.Count, 0) or 0),
                                          "reason": "time budget used up by earlier accounts"})
            if scan_python and not deadline.expired():
                limit = _SENDER_SCAN_LIMIT if legacy else _SENT_SCAN_CAP
                in_window = int(_safe_get(lambda: items.Count, 0) or 0)
                scan_started = time.monotonic()
                stop_at = float("inf") if legacy else scan_started + min(_SENT_SCAN_SECONDS, _SENT_SCAN_TOTAL_S - scan_spent)
                matched = 0
                scanned = 0
                stopped = ""
                for item in _iter_items(items, limit):
                    if deadline.expired():
                        partial = True
                        stopped = "call time limit"
                        break
                    if time.monotonic() >= stop_at:
                        stopped = "time cap"
                        break
                    scanned += 1
                    try:
                        if int(item.Class) != _OL_CLASS_MAIL:
                            continue
                        hay = " ".join(str(_safe_get(lambda a=a: getattr(item, a)) or "") for a in ("To", "CC", "BCC")).lower()
                        hit = needle_lower in hay
                        if not hit:
                            for rec in format_recipients(item):
                                if needle_lower in (rec.get("address") or "") or needle_lower in (rec.get("name") or "").lower():
                                    hit = True
                                    break
                        if not hit:
                            continue
                        sent_here.append(_sent_summary(item, name))
                        matched += 1
                        if matched >= count:
                            break
                    except Exception:
                        continue
                if not stopped and scanned >= limit and in_window > scanned:
                    stopped = "message cap"
                # Only a real gap is reported: the walk stopped with messages still unread in
                # the window and fewer matches than asked for. A walk that reached the end of
                # the window, or found enough, is complete.
                if stopped and matched < count:
                    sent_scan_partial.append({"account": name, "scanned": scanned, "of": in_window, "reason": stopped})
                scan_spent += time.monotonic() - scan_started
            seen_ids = {str(r.get("entry_id")) for r in sent}
            sent.extend(r for r in sent_here if str(r.get("entry_id")) not in seen_ids)
        received.sort(key=lambda r: r.get("received_time") or "", reverse=True)
        sent.sort(key=lambda r: r.get("sent_time") or r.get("received_time") or "", reverse=True)
        copied.sort(key=lambda r: r.get("received_time") or "", reverse=True)
        timeline = sorted(received + sent + copied,
                          key=lambda r: r.get("sent_time") or r.get("received_time") or "", reverse=True)[:count]

        # Nothing from or to this name as a person? It may be a company, deal or topic
        # (the wrong tool for it, but a common choice). One cheap subject pass over the
        # same folders plus Sent Items, so the caller still gets the thread.
        subject_matches: list = []
        if not legacy and not received and not sent and not copied and not deadline.expired():
            for store in stores:
                if deadline.expired():
                    partial = True
                    break
                name = str(store.DisplayName)
                folders = list(inbox_trees.get(name, []))
                try:
                    folders.append(store.GetDefaultFolder(OL_FOLDER_SENT_MAIL))
                except Exception:
                    pass
                found, cut = _collect_summaries(
                    folders, dasl.sql(dasl.subject_filter(needle), dasl.received_since(since)), count, "")
                partial = partial or cut
                for r in found:
                    r["account"] = name
                subject_matches.extend(found)
            subject_matches.sort(key=lambda r: r.get("received_time") or "", reverse=True)
            subject_matches = subject_matches[:count]

        meetings = None
        if include_meetings and not deadline.expired():
            past: list = []
            upcoming: list = []
            horizon = now + timedelta(days=90)
            for store in stores:
                if deadline.expired():
                    partial = True
                    break
                try:
                    calendars = resolve_calendars(store, "all")
                except Exception:
                    continue
                if legacy:
                    events = [ev for ev in _events_in(calendars, since, horizon, 500) if not ev.get("error")]
                    events = [ev for ev in events if needle_lower in " ".join(
                        str(ev.get(k) or "") for k in ("organizer", "required_attendees", "optional_attendees")).lower()]
                else:
                    calendars = [(p, f) for p, f in calendars if is_people_calendar(p)]
                    events, cut = _meetings_with(calendars, since, horizon, needle_lower)
                    partial = partial or cut
                for ev in events:
                    ev["account"] = str(store.DisplayName)
                    (upcoming if (ev.get("start") or "") >= now.isoformat(timespec="minutes") else past).append(ev)
            past.sort(key=lambda e: e.get("start") or "")
            upcoming.sort(key=lambda e: e.get("start") or "")
            meetings = {"last": past[-1] if past else None, "next": upcoming[0] if upcoming else None,
                        "past_count": len(past), "upcoming_count": len(upcoming)}
        elif include_meetings:
            partial = True

        out = {
            "person": needle,
            "window_days": days,
            "accounts_searched": searched,
            "last_received": received[0] if received else None,
            "last_sent": sent[0] if sent else None,
            "received_count": len(received),
            "sent_count": len(sent),
            "copied_count": len(copied),
            "counts_note": (f"Counts are capped at about {count} per direction per account; raise count for more. "
                            "copied = someone else wrote to this person and the user was on the message."),
            "recent": timeline,
            "meetings": meetings,
            "elapsed_s": round(time.perf_counter() - t0, 1),
        }
        if subject_matches:
            out["subject_matches"] = subject_matches
            out["subject_matches_note"] = (
                f"No mail from or to {needle!r} as a person. These messages carry it in the subject, so it "
                "is probably a company, deal or topic: use search_emails (query=, folder=\"all\") for those, "
                "and get_conversation on a hit for the thread.")
        if sent_scan_partial:
            out["sent_scan_capped"] = True
            out["sent_scan_partial"] = sent_scan_partial
            where = "; ".join(f"{p['account']}: {p['scanned']} of {p['of']} sent messages ({p['reason']})"
                              for p in sent_scan_partial)
            out["sent_scan_note"] = (
                "Sent Items were only partly checked for recipients Outlook shows by name rather than address, "
                f"in {where}. Mail sent to this address under a display name there may be missing. Name the "
                "person as Outlook shows them to search Sent Items natively.")
        if skipped:
            out["accounts_skipped"] = skipped
        if partial:
            out["partial"] = True
            out["note"] = deadline.partial_note("The sweep")
        return json.dumps(out, indent=2, default=str)

    try:
        return await bridge.call(_history, person, days, count, account, include_meetings)
    except Exception as e:
        return f"Error building contact history: {format_com_error(e)}"


# =====================================================================
# TOOL: update_draft
# =====================================================================

@mcp.tool()
async def update_draft(
    entry_id: str,
    body: str = "",
    html_body: str = "",
    subject: str = "",
    to: str = "",
    cc: str = "",
    bcc: str = "",
    account: str = "",
    replace_body: bool = False,
    inline_replies: str = "",
) -> str:
    """Change an existing draft instead of creating a new one. It is NOT sent.

    BEFORE CALLING THIS: show the revised draft in the chat first and ask
    explicitly whether to apply it to the draft in Outlook; call this tool only
    after the user says yes.

    The server marks the note it writes into a draft, so a new body replaces
    just that note and leaves the signature and any quoted thread underneath
    untouched. Subject and recipients can be changed independently. Works on
    drafts in the Drafts folder (new emails, replies, forwards).

    Args:
        entry_id: EntryID of the draft (returned when it was created, or from
            list_emails with folder="drafts").
        body: New note text, plain prose with blank-line paragraphs. Leave empty
            to keep the current text.
        html_body: New note as HTML instead of body.
        subject: New subject. Empty = unchanged.
        to, cc, bcc: New recipient lists (semicolons). Empty = unchanged; each
            one given replaces that whole list.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.
        replace_body: If the draft has no marked note (it was written in Outlook,
            or edited there in a way that lost the marker), the body cannot be
            replaced safely and the tool says so. Pass true to overwrite the
            whole body anyway (signature and quoted text are lost).
        inline_replies: Optional. The complete new set of inline answers, in the
            form draft_reply takes; the ones written earlier are removed first.
            Empty = unchanged.

    Returns:
        JSON with status "draft_saved", sent=false, entry_id, folder and the
        list of fields changed.
    """
    def _update(outlook, namespace, entry_id, body, html_body, subject, to, cc, bcc, account, replace_body,
                inline_replies):
        try:
            answers = _parse_inline_replies(inline_replies)
        except ValueError as e:
            return json.dumps({"error": str(e)})
        if answers and plain_text_mode():
            return json.dumps({"error": "inline_replies needs HTML drafts; the plain_text setting is on."})
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_MAIL, "mail item"):
            return err
        if bool(_safe_get(lambda: item.Sent, False)):
            return json.dumps({"error": "That message has already been sent; only drafts can be updated."})
        changed = []
        if answers:
            missing = _write_inline_answers(item, answers)
            if missing:
                return _inline_error(missing)
            changed.append("inline answers")
        if subject:
            item.Subject = subject
            changed.append("subject")
        if to:
            item.To = to
            changed.append("to")
        if cc:
            item.CC = cc
            changed.append("cc")
        if bcc:
            item.BCC = bcc
            changed.append("bcc")
        if body or html_body:
            if plain_text_mode():
                item.Body = body or html_body
                changed.append("body (plain text, whole body replaced)")
            else:
                fragment = style_html(html_body) if html_body else text_to_html(body)
                current = str(_safe_get(lambda: item.HTMLBody) or "")
                new_html = replace_note(current, fragment)
                if new_html is None:
                    if not replace_body:
                        return json.dumps({
                            "error": "This draft has no marked note section to replace (it was not written by this "
                                     "server, or it was edited in Outlook and the marker was lost). Pass "
                                     "replace_body=true to overwrite the whole body, or make the edit in Outlook.",
                            "entry_id": entry_id,
                        })
                    new_html = wrap_note(fragment)
                    changed.append("body (whole body replaced)")
                else:
                    changed.append("body")
                item.HTMLBody = new_html
        if not changed:
            return json.dumps({"error": "Nothing to change: give a body, subject or recipients."})
        return _save_draft(item, "update", changed=changed,
                           subject=str(_safe_get(lambda: item.Subject) or ""),
                           to=str(_safe_get(lambda: item.To) or ""), cc=str(_safe_get(lambda: item.CC) or ""))

    try:
        return await bridge.call(_update, entry_id, body, html_body, subject, to, cc, bcc, account, replace_body,
                                 inline_replies)
    except Exception as e:
        return f"Error updating draft: {format_com_error(e)}"


# =====================================================================
# TOOL: add_attachment
# =====================================================================

@mcp.tool()
async def add_attachment(entry_id: str, paths: str, account: str = "") -> str:
    """Attach one or more files from this PC to an existing draft (email, reply,
    forward or unsent meeting). Nothing is sent.

    Args:
        entry_id: EntryID of the draft.
        paths: Full path(s) of the file(s) on this computer, separated by
            semicolons, e.g. "C:\\Users\\me\\Documents\\Model.xlsx".
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        JSON with the files attached (filename, size) and the draft's entry_id.
    """
    def _add(outlook, namespace, entry_id, paths, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if bool(_safe_get(lambda: item.Sent, False)):
            return json.dumps({"error": "That message has already been sent; files can only be added to drafts."})
        added = _attach_files(item, paths)
        item.Save()
        return json.dumps({"status": "attached", "entry_id": str(_safe_get(lambda: item.EntryID) or entry_id),
                           "files_attached": added, "attachment_count": int(_safe_get(lambda: item.Attachments.Count, 0) or 0)},
                          indent=2, default=str)

    try:
        return await bridge.call(_add, entry_id, paths, account)
    except Exception as e:
        return f"Error adding attachment: {format_com_error(e)}"


# =====================================================================
# TOOL: health_check
# =====================================================================

@mcp.tool()
async def health_check() -> str:
    """Report whether the connector is working and why not if it isn't: server
    version, Outlook running and its version, accounts loaded and which one is
    the default for tools, current settings, and whether Outlook is blocking
    address access behind its "A program is trying to access..." prompt.

    Use this first when something fails, especially from a phone where the
    laptop screen cannot be seen.

    body_search_index and legacy_mode are diagnostics for the slow-search
    troubleshooting in the README. Do not mention them unless the user asks
    about search speed or a search was slow: "not indexed" is the normal
    state on many PCs and only affects include_body searches.

    Returns:
        JSON with status ("ok" or "degraded"), the checks above, and a
        plain-language "warning" when something needs the user's attention.
    """
    result: dict = {
        "server_version": __version__,
        "python": sys.version.split()[0],
        "platform": f"{platform.system()} {platform.machine()}",
        "settings_file": str(settings.settings_path()),
        "settings": {k: v["value"] for k, v in settings.all_settings().items()},
    }
    warnings: list[str] = []

    def _basic(outlook, namespace):
        stores = []
        default_id = str(namespace.DefaultStore.StoreID)
        include_shared = bool(settings.get("include_shared_in_all"))
        for i in range(namespace.Stores.Count):
            store = namespace.Stores.Item(i + 1)
            hidden = _is_hidden_store(store)
            shared = _is_shared_uncached(store)
            stores.append({
                "display_name": str(store.DisplayName),
                "is_outlook_default": str(store.StoreID) == default_id,
                "hidden": hidden,
                "shared_uncached": shared,
                "in_all_sweeps": (not hidden) and (include_shared or not shared),
            })
        tools_default = _resolve_store(namespace, "")
        _body_index_ok(namespace)
        return {
            "outlook_running": True,
            "outlook_version": str(_safe_get(lambda: outlook.Version) or ""),
            "accounts": stores,
            "default_account_for_tools": str(tools_default.DisplayName) if tools_default is not None else None,
            "body_search_index": _BODY_INDEX["detail"],
            "legacy_mode": _legacy(),
        }

    try:
        result.update(await bridge.call(_basic))
    except Exception as e:
        result["outlook_running"] = False
        result["error"] = format_com_error(e)
        warnings.append("Outlook is not reachable. Make sure classic Outlook is open (the 'New Outlook' switch off) and try again.")

    if result.get("outlook_running"):
        def _probe(outlook, namespace):
            store = _resolve_store(namespace, "")
            items = store.GetDefaultFolder(OL_FOLDER_INBOX).Items
            items.Sort("[ReceivedTime]", True)
            item = items.GetFirst()
            if item is not None:
                _ = str(_safe_get(lambda: item.SenderEmailAddress) or "")
                _ = get_smtp(item)
            return True

        try:
            await bridge.call(_probe, _timeout=4)
            result["address_access"] = "ok"
        except TimeoutError:
            result["address_access"] = "blocked"
            warnings.append("Outlook is probably showing its 'A program is trying to access email address information' "
                            "prompt. On the laptop, click 'Allow access for 10 minutes'. See README for permanent options.")
        except Exception as e:
            result["address_access"] = f"error: {format_com_error(e)}"

    preferred = settings.get("default_account")
    if preferred and result.get("accounts") and not any(preferred.lower() in a["display_name"].lower() for a in result["accounts"]):
        warnings.append(f"default_account is set to {preferred!r} but no loaded account matches it; Outlook's own default is being used.")

    result["status"] = "ok" if not warnings else "degraded"
    if warnings:
        result["warning"] = " ".join(warnings)
    return json.dumps(result, indent=2, default=str)



def main():
    logger.info("Starting Outlook Desktop MCP server %s...", __version__)
    bridge.start()
    logger.info("COM bridge ready. Starting MCP stdio transport...")
    try:
        mcp.run(transport="stdio")
    finally:
        bridge.stop()


if __name__ == "__main__":
    main()
