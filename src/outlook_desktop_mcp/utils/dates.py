"""Outlook dates → ISO 8601 with the machine's UTC offset.

COM hands back ``pywintypes.datetime`` values in local time; ``str()`` on them gives an
inconsistent, timezone-less format. Everything the server returns goes through ``iso``
so a model can compare and reason about times. Outlook represents "no date" as the year
4501, which becomes ``None``.
"""
from datetime import datetime

OUTLOOK_NO_DATE_YEAR = 4500


def to_local(value) -> datetime | None:
    """Naive local ``datetime`` from a COM/pywintypes date, or None for "no date"."""
    if value is None:
        return None
    try:
        if value.year >= OUTLOOK_NO_DATE_YEAR:
            return None
        return datetime(value.year, value.month, value.day, value.hour, value.minute, value.second)
    except Exception:
        return None


def iso(value) -> str | None:
    """``2026-10-26T17:40:00-04:00`` or None."""
    local = to_local(value)
    if local is None:
        return None
    return local.astimezone().isoformat(timespec="seconds")


def iso_date(value) -> str | None:
    """``2026-10-26`` or None (for due dates and all-day events)."""
    local = to_local(value)
    return local.date().isoformat() if local else None
