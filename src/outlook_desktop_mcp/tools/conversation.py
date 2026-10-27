"""Reconstruct the whole conversation a message belongs to.

Walks Outlook's conversation tree (``GetConversation`` → ``GetRootItems`` →
``GetChildren``). Those return ``SimpleItems``, which only support ``Count``/``Item(i)``
— ``GetFirst/GetNext`` silently return nothing on them. When the tree is empty or
unavailable (items moved between folders, non-Exchange stores) a ``ConversationTopic``
search over the store's Inbox subtree and Sent Items is used instead, kept to the same
``ConversationID``. Messages are deduped by Internet Message-ID so a copy in Sent Items
and a copy in the Inbox count once.
"""
from outlook_desktop_mcp.tools._folder_constants import OL_FOLDER_INBOX, OL_FOLDER_SENT_MAIL
from outlook_desktop_mcp.utils.dates import to_local

PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
MAIL_CLASS = 43
MAX_NODES = 500
MAX_TOPIC_MATCHES = 200
MAX_DEPTH = 50
_EXCLUDED_FOLDERS = {"deleted items", "junk email", "junk e-mail", "drafts", "outbox", "conversation history"}


def items_by_index(collection, limit: int = MAX_NODES) -> list:
    out = []
    try:
        for i in range(min(int(collection.Count), limit)):
            out.append(collection.Item(i + 1))
    except Exception:
        pass
    return out


def folder_path(item) -> str:
    """``Inbox/Clients`` from the item's parent ``FolderPath`` (``\\\\Store\\Inbox\\Clients``)."""
    try:
        raw = str(item.Parent.FolderPath).lstrip("\\")
        parts = [p for p in raw.split("\\") if p]
        return "/".join(parts[1:]) if len(parts) > 1 else (parts[0] if parts else "")
    except Exception:
        return ""


def message_id(item) -> str:
    try:
        return str(item.PropertyAccessor.GetProperty(PR_INTERNET_MESSAGE_ID) or "").strip()
    except Exception:
        return ""


def walk_conversation(item) -> list | None:
    """Mail items reachable from ``item.GetConversation()``; None when unavailable."""
    try:
        conv = item.GetConversation()
    except Exception:
        return None
    if conv is None:
        return None
    found: list = []

    def visit(node, depth: int) -> None:
        if len(found) >= MAX_NODES or depth > MAX_DEPTH:
            return
        try:
            if int(node.Class) == MAIL_CLASS:
                found.append(node)
        except Exception:
            pass
        try:
            children = conv.GetChildren(node)
        except Exception:
            return
        for child in items_by_index(children):
            visit(child, depth + 1)

    try:
        roots = items_by_index(conv.GetRootItems())
    except Exception:
        return None
    for root in roots:
        visit(root, 0)
    return found


def _subtree(folder, out: list, depth: int = 0) -> None:
    out.append(folder)
    if depth > 10:
        return
    try:
        for i in range(folder.Folders.Count):
            sub = folder.Folders.Item(i + 1)
            if str(sub.Name or "").lower() in _EXCLUDED_FOLDERS:
                continue
            _subtree(sub, out, depth + 1)
    except Exception:
        pass


def search_by_topic(store, item) -> list:
    """Fallback: ``[ConversationTopic]`` Restrict over Inbox subtree + Sent Items, same ConversationID."""
    try:
        topic = str(item.ConversationTopic or "").strip()
        conv_id = str(item.ConversationID or "")
    except Exception:
        return []
    if not topic:
        return []
    folders: list = []
    for folder_id in (OL_FOLDER_INBOX, OL_FOLDER_SENT_MAIL):
        try:
            _subtree(store.GetDefaultFolder(folder_id), folders)
        except Exception:
            continue
    filt = "[ConversationTopic] = '" + topic.replace("'", "''") + "'"
    found = []
    for folder in folders:
        try:
            matches = folder.Items.Restrict(filt)
            node = matches.GetFirst()
            while node and len(found) < MAX_TOPIC_MATCHES:
                try:
                    if int(node.Class) == MAIL_CLASS and (not conv_id or str(node.ConversationID) == conv_id):
                        found.append(node)
                except Exception:
                    pass
                node = matches.GetNext()
        except Exception:
            continue
    return found


def _when(node):
    for attr in ("SentOn", "ReceivedTime"):
        try:
            value = to_local(getattr(node, attr))
            if value is not None:
                return value
        except Exception:
            continue
    return None


def conversation_items(store, item, limit: int = 30) -> tuple[list, str]:
    """``(mail items oldest → newest including ``item``, source)``.

    ``source`` is "conversation" (tree), "topic_search" (fallback) or "none" (only the
    item itself). Items in Deleted Items, Junk, Drafts and Outbox are left out. When more
    than ``limit`` messages exist the oldest are dropped.
    """
    nodes = walk_conversation(item)
    source = "conversation"
    if not nodes or len(nodes) <= 1:
        nodes = search_by_topic(store, item)
        source = "topic_search" if len(nodes) > 1 else "none"
    seen: set[str] = set()
    kept = []
    for node in [item, *(nodes or [])]:
        try:
            key = message_id(node) or str(node.EntryID)
        except Exception:
            continue
        if key in seen:
            continue
        path = folder_path(node)
        if any(part.lower() in _EXCLUDED_FOLDERS for part in path.split("/")) and node is not item:
            continue
        seen.add(key)
        kept.append((_when(node), path, node))
    kept.sort(key=lambda entry: (entry[0] is None, entry[0] or 0))
    if len(kept) > limit:
        kept = kept[-limit:]
    return [(path, node) for _when_, path, node in kept], source
