"""Read-only mode: hide every tool that can change mailbox state.

Enabled with ``outlook-desktop-mcp --read-only`` or the environment variable
``OUTLOOK_MCP_READ_ONLY=1``. In this mode the server registers only tools that read
(list, search, get, save attachments to disk) and never registers send, reply, move,
mark, create, update, delete, respond, complete or toggle tools — so an MCP client
cannot act on the mailbox even if a prompt (or a malicious email) asks it to.

The decision is made at import time because FastMCP registers tools through decorators.
"""
import os

ENV_VAR = "OUTLOOK_MCP_READ_ONLY"
_TRUE = {"1", "true", "yes", "on"}


def is_read_only() -> bool:
    return os.environ.get(ENV_VAR, "").strip().lower() in _TRUE


def enable_read_only() -> None:
    os.environ[ENV_VAR] = "1"


# Tools that modify mailbox state. Kept here so tests and docs have one list to check.
WRITE_TOOLS = frozenset({
    "send_email", "reply_email", "move_email", "mark_as_read", "mark_as_unread",
    "create_event", "create_meeting", "update_event", "delete_event", "respond_to_meeting",
    "create_task", "complete_task", "delete_task",
    "set_category", "toggle_rule",
})
