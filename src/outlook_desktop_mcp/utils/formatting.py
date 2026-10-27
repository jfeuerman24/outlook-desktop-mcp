"""Helpers for extracting and formatting Outlook item data.

Dates are ISO 8601 with the local UTC offset (``utils.dates``), sender addresses are
resolved to SMTP (``utils.addresses``), and full bodies have their quoted history removed
before the length cap (``utils.text``) unless the caller asks for it.
"""
import re

from outlook_desktop_mcp.tools._folder_constants import (
    BUSY_STATUS_NAMES,
    MEETING_STATUS_NAMES,
    RESPONSE_NAMES,
    TASK_STATUS_NAMES,
    IMPORTANCE_NAMES,
)
from outlook_desktop_mcp.tools.conversation import thread_depth
from outlook_desktop_mcp.utils.addresses import format_recipients, get_smtp
from outlook_desktop_mcp.utils.dates import iso, iso_date
from outlook_desktop_mcp.utils.text import clean_body


def truncate(text: str, max_length: int = 2000) -> str:
    if len(text) <= max_length:
        return text
    return text[:max_length] + "\n... [truncated]"


def strip_html(html: str) -> str:
    text = re.sub(r"<[^>]+>", "", html)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _folder_path(item) -> str:
    try:
        raw = str(item.Parent.FolderPath).lstrip("\\")
        parts = [p for p in raw.split("\\") if p]
        return "/".join(parts[1:]) if len(parts) > 1 else (parts[0] if parts else "")
    except Exception:
        return ""


def _safe(getter, default=""):
    try:
        value = getter()
        return default if value is None else value
    except Exception:
        return default


def format_email_summary(item) -> dict:
    """Extract key fields from an Outlook MailItem into a dict."""
    return {
        "entry_id": item.EntryID,
        "subject": item.Subject or "(no subject)",
        "sender": get_smtp(item) or str(_safe(lambda: item.SenderEmailAddress, "unknown")),
        "sender_name": str(_safe(lambda: item.SenderName, "unknown")),
        "received_time": iso(_safe(lambda: item.ReceivedTime, None)),
        "folder": _folder_path(item),
        "unread": bool(item.UnRead),
        "has_attachments": bool(item.Attachments.Count > 0),
        "attachment_count": item.Attachments.Count,
        "conversation_id": str(_safe(lambda: item.ConversationID)),
        "thread_depth": thread_depth(_safe(lambda: item.ConversationIndex, None)),
        "categories": str(_safe(lambda: item.Categories)),
    }


def format_email_full(item, body_max_length: int = 5000, trim_history: bool = True) -> dict:
    """Extract full email details including body.

    ``trim_history=True`` (default) removes the quoted earlier messages under a reply
    before applying the length cap; ``body_trimmed`` says whether anything was removed.
    """
    result = format_email_summary(item)
    result["sent_time"] = iso(_safe(lambda: item.SentOn, None))
    result["to"] = item.To or ""
    result["cc"] = item.CC or ""
    result["recipients"] = format_recipients(item)
    result.update(clean_body(item.Body or "", trim_history=trim_history, max_length=body_max_length))
    return result


# --- Calendar formatting ---


def format_event_summary(item, calendar: str = "") -> dict:
    """Extract key fields from an Outlook AppointmentItem."""
    summary = {
        "entry_id": item.EntryID,
        "subject": item.Subject or "(no subject)",
        "start": iso(item.Start),
        "end": iso(item.End),
        "duration": item.Duration,
        "location": item.Location or "",
        "organizer": item.Organizer or "",
        "is_recurring": bool(item.IsRecurring),
        "all_day": bool(item.AllDayEvent),
        "busy_status": BUSY_STATUS_NAMES.get(item.BusyStatus, "unknown"),
        "meeting_status": MEETING_STATUS_NAMES.get(item.MeetingStatus, "unknown"),
        "required_attendees": item.RequiredAttendees or "",
        "optional_attendees": item.OptionalAttendees or "",
        "categories": str(_safe(lambda: item.Categories)),
    }
    if calendar:
        summary["calendar"] = calendar
    return summary


def format_event_full(item, body_max_length: int = 5000, calendar: str = "") -> dict:
    """Full event details including body."""
    result = format_event_summary(item, calendar)
    result["body"] = truncate(item.Body or "", body_max_length)
    result["reminder_set"] = bool(item.ReminderSet)
    result["reminder_minutes"] = (
        item.ReminderMinutesBeforeStart if item.ReminderSet else None
    )
    result["categories"] = item.Categories or ""
    result["response_status"] = RESPONSE_NAMES.get(item.ResponseStatus, "unknown")
    return result


# --- Task formatting ---


def format_task_summary(item) -> dict:
    """Extract key fields from an Outlook TaskItem."""
    return {
        "entry_id": item.EntryID,
        "subject": item.Subject or "(no subject)",
        "status": TASK_STATUS_NAMES.get(item.Status, "unknown"),
        "percent_complete": item.PercentComplete,
        "due_date": iso_date(_safe(lambda: item.DueDate, None)),
        "start_date": iso_date(_safe(lambda: item.StartDate, None)),
        "importance": IMPORTANCE_NAMES.get(item.Importance, "normal"),
        "complete": bool(item.Complete),
        "categories": item.Categories or "",
        "owner": item.Owner or "",
    }


def format_task_full(item, body_max_length: int = 5000) -> dict:
    """Full task details including body."""
    result = format_task_summary(item)
    result["body"] = truncate(item.Body or "", body_max_length)
    result["reminder_set"] = bool(item.ReminderSet)
    result["date_completed"] = (
        iso(_safe(lambda: item.DateCompleted, None)) if item.Complete else None
    )
    return result
