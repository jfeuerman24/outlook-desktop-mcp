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

from outlook_desktop_mcp.tools.calendars import describe_calendars, resolve_calendars
from outlook_desktop_mcp.tools.conversation import conversation_items
from outlook_desktop_mcp.utils.addresses import get_smtp, sender_matches
from outlook_desktop_mcp.utils.html import inject_into_body, plain_text_mode, style_html, text_to_html
from outlook_desktop_mcp.tools._folder_constants import (
    FOLDER_NAME_TO_ENUM,
    OL_MAIL_ITEM,
    OL_APPOINTMENT_ITEM,
    OL_FOLDER_CALENDAR,
    OL_FOLDER_DRAFTS,
    OL_FOLDER_TASKS,
    OL_MEETING,
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
    """Sanitize a string for use in a DASL LIKE filter value.
    Escapes SQL wildcards (% and _) so user input is treated as literals,
    then escapes quote characters required by DASL syntax.
    """
    query = query.replace("%", "[%]").replace("_", "[_]")
    return query.replace("'", "''").replace('"', '""')


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
    "NOTHING THIS SERVER DOES SENDS A MESSAGE. draft_email, draft_reply and "
    "draft_meeting save the item (to Drafts, or to the calendar with invitations "
    "unsent) and by default open it in an Outlook window; the user reads it and "
    "presses Send themselves. There is no tool to send, to respond to an "
    "invitation, or to cancel a meeting the user organised, because each of those "
    "would send. Tell the user where the draft is and that they need to send it.\n\n"
    + _INSTRUCTIONS_COMMON
    + "AVAILABLE TOOL CATEGORIES:\n"
    "- Email: list, read, search (with sender filter), get_conversation, "
    "draft_email, draft_reply, mark read/unread, move, attachments\n"
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

def _resolve_store(namespace, account: str = ""):
    """Resolve an account name to an Outlook Store object.

    If account is empty, returns DefaultStore.
    Otherwise does a case-insensitive substring match on Store.DisplayName.
    """
    if not account:
        return namespace.DefaultStore

    account_lower = account.lower().strip()
    for i in range(namespace.Stores.Count):
        store = namespace.Stores.Item(i + 1)
        if account_lower in store.DisplayName.lower():
            return store

    return None


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
async def list_accounts() -> str:
    """List all Outlook accounts (stores) configured in the profile.

    Returns a JSON array of account objects with display_name, store_id,
    and is_default. Use the display_name (or a unique substring) as the
    'account' parameter in other tools to target a specific account.

    Returns:
        JSON array of account objects.
    """
    def _list(outlook, namespace):
        default_id = namespace.DefaultStore.StoreID
        results = []
        for i in range(namespace.Stores.Count):
            store = namespace.Stores.Item(i + 1)
            results.append({
                "display_name": store.DisplayName,
                "store_id": store.StoreID,
                "is_default": store.StoreID == default_id,
            })
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_list)
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
    fragment = style_html(html_body) if html_body else text_to_html(body)
    item.HTMLBody = inject_into_body(str(_safe_get(lambda: item.HTMLBody) or ""), fragment)


def _save_draft(item, kind: str, open_in_outlook: bool, **extra) -> str:
    """Save ``item`` without sending, optionally open it in Outlook, describe the result."""
    item.Save()
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
        "note": "Nothing has been sent. The user must open the draft in Outlook, check it and press Send.",
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
    open_in_outlook: bool = True,
) -> str:
    """Write a new email to the Drafts folder. It is NOT sent.

    The draft is saved in the chosen account's Drafts folder, with the user's
    signature below the text, and by default opened in an Outlook window so the
    user can read it and press Send themselves. The text is written as HTML in
    the user's default font (Aptos 12 pt black unless the OUTLOOK_MCP_FONT_*
    environment variables say otherwise); write the body as normal prose with
    blank lines between paragraphs.

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
        open_in_outlook: If true (default), open the draft in an Outlook window
            right away. Set false when preparing several drafts at once.

    Returns:
        JSON with status "draft_saved", sent=false, entry_id, folder and whether
        it was opened in Outlook. Tell the user the draft is waiting for them.
    """
    def _draft(outlook, namespace, to, subject, body, cc, bcc, html_body, account, open_in_outlook):
        store = _require_store(namespace, account)
        mail = store.GetDefaultFolder(OL_FOLDER_DRAFTS).Items.Add(OL_MAIL_ITEM)
        _set_sending_account(mail, outlook, store)
        mail.To = to
        if cc:
            mail.CC = cc
        if bcc:
            mail.BCC = bcc
        mail.Subject = subject
        _write_draft_text(mail, body, html_body)
        return _save_draft(mail, "email", open_in_outlook, to=to, cc=cc, bcc=bcc)

    try:
        return await bridge.call(_draft, to, subject, body, cc, bcc, html_body, account, open_in_outlook)
    except Exception as e:
        return f"Error writing draft: {format_com_error(e)}"


# =====================================================================
# TOOL 2: list_emails
# =====================================================================

def _collect_summaries(folders, restriction: str, count: int, sender: str) -> list:
    """Newest-first summaries across ``folders`` (restricted, optional sender filter)."""
    results = []
    for target in folders:
        items = target.Items
        items.Sort("[ReceivedTime]", True)
        if restriction:
            items = items.Restrict(restriction)
        scan = _SENDER_SCAN_LIMIT if sender else count
        matched = 0
        for item in _iter_items(items, scan):
            try:
                if int(item.Class) != _OL_CLASS_MAIL:
                    continue
                smtp = get_smtp(item)
                if sender and not sender_matches(item, sender, smtp):
                    continue
                results.append(format_email_summary(item))
                matched += 1
                if matched >= count:
                    break
            except Exception:
                continue
    results.sort(key=lambda r: r.get("received_time") or "", reverse=True)
    return results[:count]


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
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of email summary objects.
    """
    def _list(outlook, namespace, folder, count, unread_only, start_date, end_date,
              sender, include_subfolders, account):
        count = min(max(1, count), 200)
        store = _require_store(namespace, account)
        target = _resolve_folder(namespace, folder, store)
        if not target:
            return json.dumps({"error": f"Folder '{folder}' not found"})
        folders = []
        if include_subfolders:
            _folder_subtree(target, folders)
        else:
            folders.append(target)

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

        results = _collect_summaries(folders, " AND ".join(restrictions), count, sender)
        return json.dumps(results, indent=2, default=str)

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
        account: Optional. Account display name (or substring) the message lives
            in. Default: primary account.

    Returns:
        JSON object: {"topic", "source", "message_count", "messages": [...]}
        where source is "conversation", "topic_search" or "none".
    """
    def _get(outlook, namespace, entry_id, count, include_bodies, account):
        count = min(max(1, count), 200)
        store = _require_store(namespace, account)
        item = namespace.GetItemFromID(entry_id, store.StoreID) if account else namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_MAIL, "mail item"):
            return err
        messages, source = conversation_items(store, item, limit=count)
        out = []
        for path, node in messages:
            try:
                record = format_email_full(node) if include_bodies else format_email_summary(node)
                record["folder"] = path or record.get("folder", "")
                record["is_requested_message"] = str(node.EntryID) == str(item.EntryID)
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
        return await bridge.call(_get, entry_id, count, include_bodies, account)
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
    open_in_outlook: bool = True,
) -> str:
    """Write a reply to an email as a draft. It is NOT sent.

    Creates Outlook's normal reply (recipients filled in, quoted history and
    signature below), puts the text above it in the user's default font, saves
    it to Drafts, and by default opens it in an Outlook window so the user can
    read it and press Send themselves.

    Args:
        entry_id: The unique Outlook EntryID of the email to reply to.
        body: The reply text, as plain prose; blank lines separate paragraphs.
        reply_all: If true, address everyone on the original (sender + To/CC).
            If false (default), address only the sender.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.
        open_in_outlook: If true (default), open the draft in an Outlook window
            right away. Set false when preparing several drafts at once.

    Returns:
        JSON with status "draft_saved", sent=false, entry_id of the draft and
        the folder it is in. Tell the user the draft is waiting for them.
    """
    def _reply(outlook, namespace, entry_id, body, reply_all, account, open_in_outlook):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)
        if err := _check_item_class(item, _OL_CLASS_MAIL, "mail item"):
            return err
        reply_item = item.ReplyAll() if reply_all else item.Reply()
        _write_draft_text(reply_item, body)
        return _save_draft(
            reply_item, "reply_all" if reply_all else "reply", open_in_outlook,
            in_reply_to=str(item.Subject or ""), to=str(_safe_get(lambda: reply_item.To) or ""),
            cc=str(_safe_get(lambda: reply_item.CC) or ""),
        )

    try:
        return await bridge.call(_reply, entry_id, body, reply_all, account, open_in_outlook)
    except Exception as e:
        return f"Error writing reply draft: {format_com_error(e)}"


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

@mcp.tool()
async def search_emails(
    query: str = "",
    folder: str = "inbox",
    count: int = 10,
    start_date: str = "",
    end_date: str = "",
    sender: str = "",
    include_subfolders: bool = False,
    account: str = "",
) -> str:
    """Search for emails in Outlook using text search.

    Searches email subjects and bodies using Outlook's DASL filter, optionally
    limited to one sender. Results are sorted by received time (newest first).
    Each result includes entry_id for further operations.

    Args:
        query: The search term (case-insensitive substring match on subject or
            body). Examples: "budget report", "meeting notes". May be empty when
            'sender' is given, to list everything from that sender.
        folder: Folder to search in. Default "inbox". Supports the same names
            and slash paths as list_emails.
        count: Maximum results to return. Default 10, max 200.
        start_date: Optional. Only return emails received on or after this date.
            ISO 8601 format (e.g. "2026-03-10" or "2026-03-10 09:00").
        end_date: Optional. Only return emails received on or before this date.
            ISO 8601 format. Default: now (if start_date is provided).
        sender: Optional. Case-insensitive substring of the sender's email
            address or display name, e.g. "freestone" or "Linda Harlow".
        include_subfolders: If true, also search every subfolder of 'folder'.
            Default false.
        account: Optional. Account display name (or substring) to target.
            Default: primary account. Use list_accounts to see available accounts.

    Returns:
        JSON array of matching email summaries, or an error.
    """
    def _search(outlook, namespace, query, folder, count, start_date, end_date,
                sender, include_subfolders, account):
        count = min(max(1, count), 200)
        if not query and not sender:
            return json.dumps({"error": "Provide a query, a sender, or both"})
        store = _require_store(namespace, account)
        target = _resolve_folder(namespace, folder, store)
        if not target:
            return json.dumps({"error": f"Folder '{folder}' not found"})
        folders = []
        if include_subfolders:
            _folder_subtree(target, folders)
        else:
            folders.append(target)

        dasl_parts = []
        if query:
            safe_query = _safe_dasl(query)
            dasl_parts.append(
                f"(\"urn:schemas:httpmail:subject\" LIKE '%{safe_query}%' OR "
                f"\"urn:schemas:httpmail:textdescription\" LIKE '%{safe_query}%')"
            )
        if start_date:
            start = _parse_date(start_date)
            dasl_parts.append(
                f"\"urn:schemas:httpmail:datereceived\" >= '{start.strftime('%m/%d/%Y %H:%M')}'"
            )
        if end_date:
            end = _parse_date(end_date)
            dasl_parts.append(
                f"\"urn:schemas:httpmail:datereceived\" <= '{end.strftime('%m/%d/%Y %H:%M')}'"
            )
        elif start_date:
            dasl_parts.append(
                f"\"urn:schemas:httpmail:datereceived\" <= '{datetime.now().strftime('%m/%d/%Y %H:%M')}'"
            )

        filter_str = ("@SQL=" + " AND ".join(dasl_parts)) if dasl_parts else ""
        results = _collect_summaries(folders, filter_str, count, sender)
        return json.dumps(results, indent=2, default=str)

    try:
        return await bridge.call(_search, query, folder, count, start_date, end_date,
                                 sender, include_subfolders, account)
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
            results.append({"calendar": label, "error": format_com_error(e)})
            continue
        seen: set[tuple[str, str]] = set()
        for item in _iter_items(filtered, _EVENT_SCAN_LIMIT):
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
    open_in_outlook: bool = True,
) -> str:
    """Prepare a meeting with attendees on the calendar. Invitations are NOT sent.

    The meeting is saved on the user's calendar with the attendees filled in
    and invitations unsent (Outlook marks it "invitations have not been sent"),
    and by default opened in an Outlook window so the user can check it and
    press Send themselves.

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
        open_in_outlook: If true (default), open the meeting in an Outlook
            window right away.

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

def main():
    logger.info("Starting Outlook Desktop MCP server...")
    bridge.start()
    logger.info("COM bridge ready. Starting MCP stdio transport...")
    try:
        mcp.run(transport="stdio")
    finally:
        bridge.stop()


if __name__ == "__main__":
    main()
