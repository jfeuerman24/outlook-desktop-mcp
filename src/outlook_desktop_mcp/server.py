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

from outlook_desktop_mcp import __version__

from outlook_desktop_mcp.tools import calendars as calendars_mod
from outlook_desktop_mcp.tools.calendars import describe_calendars, is_people_calendar, resolve_calendars
from outlook_desktop_mcp.tools.conversation import conversation_items
from outlook_desktop_mcp.utils import dasl, deadline, settings
from outlook_desktop_mcp.utils.addresses import format_recipients, get_smtp, sender_matches
from outlook_desktop_mcp.utils.dates import iso
from outlook_desktop_mcp.utils.extract import extract
from outlook_desktop_mcp.utils.html import (
    inject_into_body, open_drafts_default, plain_text_mode, replace_note, style_html, text_to_html, wrap_note,
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
    "NOTHING THIS SERVER DOES SENDS A MESSAGE. draft_email, draft_reply, draft_forward "
    "and draft_meeting save the item (to Drafts, or to the calendar with invitations "
    "unsent); it opens in an Outlook window only if the open_drafts setting is on. The "
    "user reads it and presses Send themselves. There is no tool to send, to respond to "
    "an invitation, or to cancel a meeting the user organised, because each of those "
    "would send. Tell the user where the draft is and that they need to send it.\n\n"
    + _INSTRUCTIONS_COMMON
    + "FINDING MAIL QUICKLY. Outlook filters by sender, recipient, subject and date in "
    "well under a second even in a 10,000-message folder, so prefer those: search_emails "
    "with sender= (a name, address or domain) and/or a subject word, in the default "
    "account first. Widen to account=\"all\" only if that finds nothing, and say so. "
    "Body-text search (include_body=true) opens every message and can take 10 to 30 "
    "seconds per large folder: use it last, with a date range. contact_history sweeps "
    "every account and calendar: use it for one specific person after a targeted search, "
    "not as the first move, and give it a full name, address or domain rather than a "
    "short first name. A result with \"partial\": true stopped at the time limit; narrow "
    "the request instead of repeating it.\n\n"
    "AVAILABLE TOOL CATEGORIES:\n"
    "- Email: list, read, search (sender/subject/body filters), get_conversation, "
    "contact_history, draft_email, draft_reply, draft_forward, update_draft, "
    "mark read/unread, move, attachments\n"
    "- Calendar: list_calendars, list events, create appointments, draft_meeting "
    "(invitations unsent), update, delete (not organised meetings), search events\n"
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


def _save_draft(item, kind: str, open_in_outlook: bool | None, **extra) -> str:
    """Save ``item`` without sending, optionally open it in Outlook, describe the result.

    ``open_in_outlook=None`` means "use the OUTLOOK_MCP_OPEN_DRAFTS setting", which is
    off by default: the draft just lands in Drafts.
    """
    item.Save()
    if open_in_outlook is None:
        open_in_outlook = open_drafts_default()
    opened = False
    if open_in_outlook:
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
    return json.dumps({
        "status": "draft_saved",
        "kind": kind,
        "sent": False,
        "subject": str(_safe_get(lambda: item.Subject) or ""),
        "entry_id": str(_safe_get(lambda: item.EntryID) or ""),
        "folder": folder,
        "opened_in_outlook": opened,
        "note": ("Nothing has been sent. " + ("It is open in an Outlook window. " if opened else "It is in the Drafts folder, not opened. ")
                 + "The user checks it in Outlook and presses Send."),
        **extra,
    }, indent=2, default=str)


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
    open_in_outlook: bool | None = None,
) -> str:
    """Write a new email to the Drafts folder. It is NOT sent.

    BEFORE CALLING THIS: show the complete draft in the chat first (recipients,
    subject, full text) and work on it there until the user is happy. Then ask
    explicitly, as its own question, whether to put it in Outlook, and call this
    tool only after the user says yes. Never call it as the first step of a
    "write an email" request. The user wants to see and edit drafts in the
    conversation, not discover them in Outlook.

    The draft is saved in the chosen account's Drafts folder, with the user's
    signature below the text, for the user to open, read and press Send
    themselves (it opens in an Outlook window only if the open_drafts setting
    or open_in_outlook says so). The text is written as HTML in
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
        open_in_outlook: Optional. True also opens the draft in an Outlook
            window; false only saves it to Drafts. Leave unset to follow the
            OUTLOOK_MCP_OPEN_DRAFTS setting (default: only save).

    Returns:
        JSON with status "draft_saved", sent=false, entry_id, folder and whether
        it was opened in Outlook. Tell the user the draft is waiting for them.
    """
    def _draft(outlook, namespace, to, subject, body, cc, bcc, html_body, account, attachments, open_in_outlook):
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
        return _save_draft(mail, "email", open_in_outlook, to=to, cc=cc, bcc=bcc, files_attached=files)

    try:
        return await bridge.call(_draft, to, subject, body, cc, bcc, html_body, account, attachments, open_in_outlook)
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
    account when more than one store is searched. Returns a JSON error string when a
    single named store lacks the folder; with several stores, stores without it are
    skipped. ``partial`` is true when the time limit stopped the search early."""
    merged = []
    partial = False
    for store in stores:
        if deadline.expired():
            partial = True
            break
        target = _resolve_folder(namespace, folder, store)
        if not target:
            if len(stores) == 1:
                return json.dumps({"error": f"Folder '{folder}' not found"}), False
            continue
        folders = []
        if include_subfolders:
            _folder_subtree(target, folders)
        else:
            folders.append(target)
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


def _results_json(results: list, partial: bool, skipped: list) -> str:
    """A plain JSON array when the search completed over every account it was asked
    for; otherwise an object with the array under "results" plus "partial" and/or
    "accounts_skipped", so the caller knows what it did not see."""
    if not partial and not skipped:
        return json.dumps(results, indent=2, default=str)
    out: dict = {"results": results}
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
) -> str:
    """List recent emails from a specified Outlook folder.

    Returns a JSON array of email summaries sorted by received time (newest
    first). Each summary includes entry_id, subject, sender (SMTP address),
    sender_name, received_time (ISO 8601), folder, unread status, attachment
    info and conversation_id.

    Use the entry_id from results to read full content with read_email,
    get_conversation for the whole thread, or actions like mark_as_read,
    move_email, or draft_reply.

    Args:
        folder: The folder to list. Case-insensitive names: "inbox" (default),
            "sent"/"sentmail", "drafts", "deleted"/"trash", "junk"/"spam",
            "outbox", "archive", a custom folder name, or a slash path such as
            "Inbox/Clients" (see list_folders).
        count: Maximum number of emails to return. Default 10, max 200.
        unread_only: If true, only return unread emails. Default false.
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
              sender, include_subfolders, account):
        count = min(max(1, count), 200)
        stores, skipped = _stores_for(namespace, account)

        # Build restriction filters
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

        results, partial = _collect_across_stores(namespace, stores, folder, include_subfolders,
                                                  " AND ".join(restrictions), count, sender)
        if isinstance(results, str):
            return results
        return _results_json(results, partial, skipped)

    try:
        return await bridge.call(_list, folder, count, unread_only, start_date, end_date,
                                 sender, include_subfolders, account)
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
        attachment info, conversation_id).
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
        }, indent=2, default=str)

    try:
        return await bridge.call(_get, entry_id, count, include_bodies, max_body_chars, account)
    except Exception as e:
        return f"Error reading conversation: {format_com_error(e)}"


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
    open_in_outlook: bool | None = None,
) -> str:
    """Write a reply to an email as a draft. It is NOT sent.

    BEFORE CALLING THIS: show the complete draft in the chat first (recipients,
    subject, full text above the quoted thread) and work on it there until the user is happy. Then ask
    explicitly, as its own question, whether to put it in Outlook, and call this
    tool only after the user says yes. Never call it as the first step of a
    "write an email" request. The user wants to see and edit drafts in the
    conversation, not discover them in Outlook.

    Creates Outlook's normal reply (recipients filled in, quoted history and
    signature below), puts the text above it in the user's default font and
    saves it to Drafts for the user to read and press Send themselves (it opens
    in an Outlook window only if the open_drafts setting or open_in_outlook
    says so).

    Args:
        entry_id: The unique Outlook EntryID of the email to reply to.
        body: The reply text, as plain prose; blank lines separate paragraphs.
        reply_all: If true, address everyone on the original (sender + To/CC).
            If false (default), address only the sender.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.
        attachments: Optional. Full path(s) of file(s) on this PC to attach,
            separated by semicolons.
        open_in_outlook: Optional. True also opens the draft in an Outlook
            window; false only saves it to Drafts. Leave unset to follow the
            OUTLOOK_MCP_OPEN_DRAFTS setting (default: only save).

    Returns:
        JSON with status "draft_saved", sent=false, entry_id of the draft and
        the folder it is in. Tell the user the draft is waiting for them.
    """
    def _reply(outlook, namespace, entry_id, body, reply_all, account, attachments, open_in_outlook):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_MAIL, "mail item"):
            return err
        reply_item = item.ReplyAll() if reply_all else item.Reply()
        files = _attach_files(reply_item, attachments)
        _write_draft_text(reply_item, body)
        return _save_draft(
            reply_item, "reply_all" if reply_all else "reply", open_in_outlook,
            in_reply_to=str(item.Subject or ""), to=str(_safe_get(lambda: reply_item.To) or ""),
            cc=str(_safe_get(lambda: reply_item.CC) or ""), files_attached=files,
        )

    try:
        return await bridge.call(_reply, entry_id, body, reply_all, account, attachments, open_in_outlook)
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
    open_in_outlook: bool | None = None,
) -> str:
    """Forward an email as a draft. It is NOT sent.

    BEFORE CALLING THIS: show the complete draft in the chat first (recipients,
    which message is being forwarded, the full note that goes above it, and
    which attachments will ride along) and work on it there until the user is
    happy. Then ask explicitly, as its own question, whether to put it in
    Outlook, and call this tool only after the user says yes. Never call it as
    the first step of a "forward this" request. The user wants to see and edit
    drafts in the conversation, not discover them in Outlook.

    Uses Outlook's own Forward(), so the "FW:" subject, the From/Sent/To/Subject
    header block, the quoted message and its attachments carry over exactly as
    a manual forward would. The note is placed above all of that in the user's
    default font, with the signature handled the same way as a reply. The
    result is saved to Drafts (opened in a window only if the
    OUTLOOK_MCP_OPEN_DRAFTS setting or open_in_outlook says so).

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
        open_in_outlook: Optional. True also opens the draft in an Outlook
            window; false only saves it to Drafts. Leave unset to follow the
            OUTLOOK_MCP_OPEN_DRAFTS setting (default: only save).

    Returns:
        JSON with status "draft_saved", sent=false, entry_id of the draft, its
        folder, opened_in_outlook, forwarded_entry_id and forwarded_subject
        (the message actually forwarded), forwarded_is_latest, and
        attachments_added (those pulled in from earlier messages). Tell the
        user where the draft is.
    """
    def _forward(outlook, namespace, entry_id, to, body, cc, bcc, account,
                 forward_latest_in_thread, include_thread_attachments, attachments, open_in_outlook):
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
                fwd, "forward", open_in_outlook,
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
            forward_latest_in_thread, include_thread_attachments, attachments, open_in_outlook,
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
            _BODY_INDEX["detail"] = "no recent message with body text to probe with"
            return False
        entry_id, word = probe
        hits = items.Restrict(dasl.sql(dasl.body_filter(word, indexed=True)))
        found = any(str(_safe_get(lambda h=h: h.EntryID) or "") == entry_id for h in _iter_items(hits, 300))
        _BODY_INDEX["ok"] = found
        _BODY_INDEX["detail"] = (f"probe word {word!r} {'found' if found else 'not found'} by the index; "
                                 + ("body searches use the index (fast)" if found else
                                    "body searches scan every message (slow); sender/subject searches are unaffected"))
        return found
    except Exception as e:
        _BODY_INDEX["detail"] = f"probe failed: {e}"
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
    account: str = "",
) -> str:
    """Find emails by sender, subject text and/or date. Fast: Outlook does the
    filtering, well under a second even in a 10,000-message folder.

    Start here for "find the email from X" or "the thread about Y". A sender
    filter alone lists everything from a person or domain; a subject word finds a
    thread. Results are newest first and each carries an entry_id for read_email,
    get_conversation or the draft tools.

    include_body=true also scans message text. That opens every message in the
    folder and takes 10 to 30 seconds per large folder; use it only when sender
    and subject searches have failed, and add a date range. (Where Windows Search
    has indexed Outlook's bodies the server detects it and body search is fast.)

    Args:
        query: Text to find in the subject (case-insensitive substring), e.g.
            "Sanders", "EHC offer". With include_body also in the message text.
            May be empty when 'sender' is given.
        folder: Folder to search in. Default "inbox". Supports the same names
            and slash paths as list_emails.
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
            Default false.
        include_body: If true, also match 'query' against message text. Slow on
            big folders (see above). Default false.
        account: Optional. Account display name (or substring), or "all" for
            every account (shared mailboxes not cached locally are skipped and
            listed under accounts_skipped). Default: the default account.

    Returns:
        JSON array of matching email summaries. If the time limit cut the search
        short or accounts were skipped, an object {"results": [...],
        "partial": true, "note": ..., "accounts_skipped": [...]} instead.
    """
    def _search(outlook, namespace, query, folder, count, start_date, end_date,
                sender, include_subfolders, include_body, account):
        count = min(max(1, count), 200)
        if not query and not sender:
            return json.dumps({"error": "Provide a query, a sender, or both"})
        stores, skipped = _stores_for(namespace, account)

        dasl_parts = []
        if query:
            if include_body or _legacy():
                indexed = (not _legacy()) and _body_index_ok(namespace)
                dasl_parts.append(dasl.subject_or_body_filter(query, indexed_body=indexed))
            else:
                dasl_parts.append(dasl.subject_filter(query))
        if start_date:
            dasl_parts.append(dasl.received_since(_parse_date(start_date)))
        if end_date:
            dasl_parts.append(dasl.received_until(_parse_date(end_date)))
        elif start_date:
            dasl_parts.append(dasl.received_until(datetime.now()))

        filter_str = dasl.sql(*dasl_parts)
        results, partial = _collect_across_stores(namespace, stores, folder, include_subfolders,
                                                  filter_str, count, sender)
        if isinstance(results, str):
            return results
        return _results_json(results, partial, skipped)

    try:
        return await bridge.call(_search, query, folder, count, start_date, end_date,
                                 sender, include_subfolders, include_body, account)
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
    """List calendar events from Outlook within a date range.

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
    account: str = "",
) -> str:
    """Create a personal calendar appointment (no attendees).

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
        location: Optional. Event location (e.g. "Conference Room A",
            "Microsoft Teams Meeting").
        body: Optional. Description or notes for the event.
        all_day: If true, creates an all-day event. Default false.
        reminder_minutes: Minutes before the event to show a reminder.
            Default 15. Set to 0 to disable reminder.
        account: Optional. Account display name (or substring) to create
            the event in. Default: primary account.

    Returns:
        Confirmation with event subject and entry_id, or an error.
    """
    def _create(outlook, namespace, subject, start, end, location, body,
                all_day, reminder_minutes, account):
        appt = outlook.CreateItem(OL_APPOINTMENT_ITEM)
        # Move to correct store's calendar if account specified
        if account:
            store = _require_store(namespace, account)
            cal = store.GetDefaultFolder(OL_FOLDER_CALENDAR)
            appt.Move(cal)
            appt = namespace.GetItemFromID(appt.EntryID)
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
        appt.Save()
        return json.dumps({
            "status": "created",
            "subject": appt.Subject,
            "start": str(appt.Start),
            "end": str(appt.End),
            "entry_id": appt.EntryID,
        }, indent=2, default=str)

    try:
        return await bridge.call(
            _create, subject, start, end, location, body, all_day,
            reminder_minutes, account,
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
    account: str = "",
    open_in_outlook: bool | None = None,
) -> str:
    """Prepare a meeting with attendees on the calendar. Invitations are NOT sent.

    BEFORE CALLING THIS: show the complete invitation in the chat first (attendees,
    subject, time, location, text) and work on it there until the user is happy. Then ask
    explicitly, as its own question, whether to put it in Outlook, and call this
    tool only after the user says yes. Never call it as the first step of a
    "write an email" request. The user wants to see and edit invitations in the
    conversation, not discover them in Outlook.

    The meeting is saved on the user's calendar with the attendees filled in
    and invitations unsent (Outlook marks it "invitations have not been sent")
    for the user to check and press Send themselves (it opens in an Outlook
    window only if the open_drafts setting or open_in_outlook says so).

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
        account: Optional. Account display name (or substring) whose calendar
            the meeting goes on and which it will be sent from. Default: primary
            account. Use list_accounts to see available accounts.
        open_in_outlook: Optional. True also opens the meeting in an Outlook
            window; false only saves it to the calendar. Leave unset to follow
            the OUTLOOK_MCP_OPEN_DRAFTS setting (default: only save).

    Returns:
        JSON with status "draft_saved", sent=false, entry_id and the calendar
        folder. Tell the user the invitation is waiting for them to send.
    """
    def _draft(outlook, namespace, subject, start, end, required_attendees,
               location, body, optional_attendees, account, open_in_outlook):
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
        return _save_draft(
            appt, "meeting", open_in_outlook,
            start=str(appt.Start), end=str(appt.End),
            required_attendees=required_attendees, optional_attendees=optional_attendees,
        )

    try:
        return await bridge.call(
            _draft, subject, start, end, required_attendees, location, body,
            optional_attendees, account, open_in_outlook,
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
    account: str = "",
) -> str:
    """Update an existing calendar event.

    Modifies properties of an appointment or meeting. Only the fields you
    provide will be updated — omitted fields remain unchanged. The change is
    saved, not sent: for a meeting the user organises, attendees are not
    notified until the user opens it in Outlook and presses Send Update.

    Args:
        entry_id: The unique Outlook EntryID of the event to update.
        subject: Optional. New event title.
        start: Optional. New start time in ISO 8601 format.
        end: Optional. New end time in ISO 8601 format.
        location: Optional. New location.
        body: Optional. New description/notes.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation with updated event details, or an error.
    """
    def _update(outlook, namespace, entry_id, subject, start, end, location, body, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
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
        item.Save()
        return json.dumps({
            "status": "updated",
            "subject": item.Subject,
            "start": str(item.Start),
            "end": str(item.End),
            "location": item.Location or "",
            "entry_id": item.EntryID,
        }, indent=2, default=str)

    try:
        return await bridge.call(
            _update, entry_id, subject, start, end, location, body, account,
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
    """Search for calendar events by keyword.

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
        JSON array of matching event summaries.
    """
    def _search(outlook, namespace, query, start_date, end_date, count, calendar, account):
        count = min(max(1, count), 200)
        store = _require_store(namespace, account)
        calendars = resolve_calendars(store, calendar)
        start = _parse_date(start_date) if start_date else datetime.now() - timedelta(days=30)
        end = _parse_date(end_date) if end_date else datetime.now() + timedelta(days=30)
        return json.dumps(_events_in(calendars, start, end, count, query), indent=2, default=str)

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
    """List tasks from the Outlook Tasks folder.

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
    """Create a new task in Outlook.

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

@mcp.tool()
async def list_categories(account: str = "") -> str:
    """List all available Outlook categories.

    Returns the color categories configured in the user's Outlook profile.
    These can be applied to emails, events, tasks, and other items.

    Args:
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of category objects with name and color index.
    """
    def _list(outlook, namespace, account):
        # Categories are profile-wide, not per-store, but we accept the param for consistency
        results = []
        for i in range(namespace.Categories.Count):
            cat = namespace.Categories.Item(i + 1)
            results.append({"name": cat.Name, "color": cat.Color})
        return json.dumps(results, indent=2, default=str)

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
    """Set categories on an email, event, or task.

    Replaces any existing categories on the item. Use comma-separated
    values for multiple categories.

    Args:
        entry_id: The EntryID of the item to categorize.
        categories: Category name(s), comma-separated. Example:
            "Important" or "Work, Follow-up". Use an empty string to
            clear all categories.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.

    Returns:
        Confirmation with the item subject and applied categories.
    """
    def _set(outlook, namespace, entry_id, categories, account):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        item.Categories = categories
        item.Save()
        return (
            f"Categories set on '{item.Subject}': "
            f"'{item.Categories or '(none)'}'"
        )

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

    Settings: default_account, hide_accounts, open_drafts, font_family,
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
        open_drafts: "true" to open each new draft in an Outlook window,
            "false" to only save it.
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
            settings.set_value(key, value)
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
    """Everything recent with one person, in one call: when they last wrote, when
    the user last wrote to them, the recent messages in both directions, and the
    last and next meetings with them. Looks across every account (unless one is
    named), the whole Inbox tree, Sent Items and every calendar.

    This is the broadest search the server offers. Use it for one specific
    person once you know who they are, not as the first move for "find the
    email about X" (use search_emails with sender= or a subject word for that).
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
        JSON with last_received, last_sent, received_count, sent_count, recent
        (merged timeline, newest first, each tagged direction and account),
        meetings {last, next, past_count, upcoming_count}, accounts_searched and
        accounts_skipped. "partial": true means the time limit stopped the sweep;
        narrow it (one account, fewer days) rather than repeating it.
    """
    def _history(outlook, namespace, person, days, count, account, include_meetings):
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
        searched = []
        partial = False
        for store in stores:
            if deadline.expired():
                partial = True
                break
            name = str(store.DisplayName)
            searched.append(name)
            # Received: Inbox tree, matched by Outlook on sender name, directory address
            # and SMTP address.
            folders: list = []
            try:
                _folder_subtree(store.GetDefaultFolder(OL_FOLDER_INBOX), folders)
            except Exception:
                folders = []
            received_filter = dasl.sql(dasl.sender_filter(needle), dasl.received_since(since))
            found, cut = _collect_summaries(folders, received_filter, count, "")
            partial = partial or cut
            for r in found:
                r["account"] = name
                r["direction"] = "received"
                received.append(r)

            # Sent: Outlook matches the To/CC/BCC display strings. Those hold names, so
            # the names this person has written under are searched alongside the needle
            # (an address needle alone would miss "Lou Castelli"). Python walks Sent
            # Items only when Outlook rejects the filter, in legacy mode, or when an
            # address-like needle found nothing natively.
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
            scan_python = items is not None and (
                legacy or not native_ok or (not sent_here and dasl.looks_like_address(needle)))
            if scan_python and not deadline.expired():
                limit = _SENDER_SCAN_LIMIT if legacy else _SENDER_SCAN_LIMIT // 2
                matched = 0
                for item in _iter_items(items, limit):
                    if deadline.expired():
                        partial = True
                        break
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
            sent.extend(sent_here)
        received.sort(key=lambda r: r.get("received_time") or "", reverse=True)
        sent.sort(key=lambda r: r.get("sent_time") or r.get("received_time") or "", reverse=True)
        timeline = sorted(received + sent, key=lambda r: r.get("sent_time") or r.get("received_time") or "", reverse=True)[:count]

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
            "counts_note": f"Counts are capped at {count} per direction per account; raise count for more.",
            "recent": timeline,
            "meetings": meetings,
        }
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
    open_in_outlook: bool | None = None,
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
        open_in_outlook: Optional. True also opens the draft in an Outlook
            window. Leave unset to follow the OUTLOOK_MCP_OPEN_DRAFTS setting.

    Returns:
        JSON with status "draft_saved", sent=false, entry_id, folder and the
        list of fields changed.
    """
    def _update(outlook, namespace, entry_id, body, html_body, subject, to, cc, bcc, account, replace_body, open_in_outlook):
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
        return _save_draft(item, "update", open_in_outlook, changed=changed,
                           subject=str(_safe_get(lambda: item.Subject) or ""),
                           to=str(_safe_get(lambda: item.To) or ""), cc=str(_safe_get(lambda: item.CC) or ""))

    try:
        return await bridge.call(_update, entry_id, body, html_body, subject, to, cc, bcc, account, replace_body, open_in_outlook)
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
