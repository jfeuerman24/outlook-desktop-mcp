"""Recurring appointments: what a tool is given, Outlook's RecurrencePattern, and back.

``parse`` turns the ``recurrence`` argument of create_event / draft_meeting into a plain
spec; ``apply`` writes that spec to an appointment's RecurrencePattern (COM) in the order
Outlook needs; ``describe`` reads an existing pattern back for get_event. Only ``parse``
and the mask helpers are pure; the tests use a fake pattern object for the other two.

Outlook quirks worth knowing: the pattern's type must be set before its days or interval;
a yearly pattern keeps Interval in months (12), so it is left alone; PatternStartDate is a
date and StartTime/EndTime carry the time of day; an occurrence of a series shares the
master's EntryID, so GetItemFromID on it returns the master. Dates go in as strings:
pywin32 shifts a naive datetime by the UTC offset on its way into COM.
"""
from __future__ import annotations

import json
import re
from datetime import date, datetime

# OlRecurrenceType
OL_RECURS_DAILY = 0
OL_RECURS_WEEKLY = 1
OL_RECURS_MONTHLY = 2
OL_RECURS_MONTH_NTH = 3
OL_RECURS_YEARLY = 5
OL_RECURS_YEAR_NTH = 6
TYPE_NAMES = {0: "daily", 1: "weekly", 2: "monthly", 3: "monthly_nth", 5: "yearly", 6: "yearly_nth"}
TYPE_VALUES = {"daily": OL_RECURS_DAILY, "weekly": OL_RECURS_WEEKLY, "monthly": OL_RECURS_MONTHLY,
               "yearly": OL_RECURS_YEARLY}

# OlDaysOfWeek bit mask
DAY_MASKS = {"sun": 1, "mon": 2, "tue": 4, "wed": 8, "thu": 16, "fri": 32, "sat": 64}
DAY_ORDER = ("mon", "tue", "wed", "thu", "fri", "sat", "sun")
WEEKDAYS_MASK = DAY_MASKS["mon"] | DAY_MASKS["tue"] | DAY_MASKS["wed"] | DAY_MASKS["thu"] | DAY_MASKS["fri"]

# OlRecurrenceState
STATE_NAMES = {0: "not_recurring", 1: "master", 2: "occurrence", 3: "exception"}

_SHORTHAND = re.compile(r"^(daily|weekly|weekdays|monthly|yearly)$", re.I)


def day_mask(days) -> int:
    """``["tue", "Thursday"]`` or ``"tue,thu"`` -> OlDaysOfWeek mask. ValueError on an unknown day."""
    if isinstance(days, str):
        days = re.split(r"[,\s;]+", days.strip())
    mask = 0
    for raw in days:
        key = str(raw or "").strip().lower()[:3]
        if not key:
            continue
        if key not in DAY_MASKS:
            raise ValueError(f"Unknown day {raw!r}; use mon, tue, wed, thu, fri, sat, sun")
        mask |= DAY_MASKS[key]
    return mask


def mask_days(mask: int) -> list[str]:
    return [d for d in DAY_ORDER if int(mask or 0) & DAY_MASKS[d]]


def _date(value) -> date | None:
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return datetime.fromisoformat(str(value).strip()).date()


def parse(text: str, start: datetime) -> dict | None:
    """The recurrence a tool was given, as a spec, or None for no recurrence.

    ``text`` is one word (``daily``, ``weekly``, ``weekdays``, ``monthly``, ``yearly``:
    weekly on the start's weekday, monthly on its day, yearly on its date, no end) or a
    JSON object: ``{"type": "weekly", "interval": 2, "days": ["tue", "thu"],
    "until": "2026-12-31"}`` or ``"count": 10`` instead of until; ``"day_of_month"`` for
    monthly. ValueError says what is wrong.
    """
    raw = (text or "").strip()
    if not raw:
        return None
    if _SHORTHAND.match(raw):
        data: dict = {"type": raw.lower()}
    else:
        try:
            data = json.loads(raw)
        except Exception as e:
            raise ValueError(f'recurrence must be one of daily, weekly, weekdays, monthly, yearly, or a JSON object '
                             f'{{"type": ..., "interval": ..., "days": [...], "until"/"count": ...}}: {e}')
        if not isinstance(data, dict):
            raise ValueError("recurrence JSON must be an object")
    kind = str(data.get("type") or "").strip().lower()
    weekdays = kind == "weekdays"
    if weekdays:
        kind = "weekly"
    if kind not in TYPE_VALUES:
        raise ValueError(f"recurrence type must be daily, weekly, weekdays, monthly or yearly; got {kind!r}")
    spec = {
        "type": TYPE_VALUES[kind],
        "interval": max(1, int(data.get("interval") or 1)),
        "day_mask": 0,
        "day_of_month": 0,
        "month_of_year": 0,
        "until": _date(data.get("until")),
        "count": int(data.get("count") or 0) or None,
    }
    if spec["until"] and spec["count"]:
        raise ValueError("recurrence takes until or count, not both")
    if spec["type"] == OL_RECURS_WEEKLY:
        if weekdays:
            spec["day_mask"] = WEEKDAYS_MASK
        elif data.get("days"):
            spec["day_mask"] = day_mask(data["days"])
        else:
            spec["day_mask"] = DAY_MASKS[DAY_ORDER[start.weekday()]]
        if not spec["day_mask"]:
            raise ValueError("a weekly recurrence needs at least one day")
    elif spec["type"] == OL_RECURS_MONTHLY:
        spec["day_of_month"] = int(data.get("day_of_month") or start.day)
    elif spec["type"] == OL_RECURS_YEARLY:
        spec["day_of_month"] = int(data.get("day_of_month") or start.day)
        spec["month_of_year"] = int(data.get("month") or start.month)
    return spec


def apply(pattern, spec: dict, start: datetime, end: datetime) -> None:
    """Write ``spec`` to an appointment's RecurrencePattern. Call after the appointment's
    Start and End are set; the appointment is saved by the caller."""
    pattern.RecurrenceType = spec["type"]
    if spec["type"] == OL_RECURS_WEEKLY:
        pattern.DayOfWeekMask = spec["day_mask"]
    elif spec["type"] == OL_RECURS_MONTHLY:
        pattern.DayOfMonth = spec["day_of_month"]
    elif spec["type"] == OL_RECURS_YEARLY:
        pattern.MonthOfYear = spec["month_of_year"]
        pattern.DayOfMonth = spec["day_of_month"]
    if spec["type"] != OL_RECURS_YEARLY:
        pattern.Interval = spec["interval"]
    pattern.PatternStartDate = start.strftime("%Y-%m-%d")
    pattern.StartTime = start.strftime("%Y-%m-%d %H:%M:%S")
    pattern.EndTime = end.strftime("%Y-%m-%d %H:%M:%S")
    if spec["until"]:
        pattern.PatternEndDate = spec["until"].strftime("%Y-%m-%d")
    elif spec["count"]:
        pattern.Occurrences = spec["count"]
    else:
        pattern.NoEndDate = True


def describe(pattern) -> dict:
    """An existing pattern as plain data: type, interval, days, day_of_month, month,
    pattern_start, and pattern_end or occurrences or no_end."""
    kind = int(pattern.RecurrenceType)
    out: dict = {"type": TYPE_NAMES.get(kind, str(kind)), "interval": int(pattern.Interval)}
    if kind in (OL_RECURS_WEEKLY, OL_RECURS_MONTH_NTH, OL_RECURS_YEAR_NTH):
        out["days"] = mask_days(int(pattern.DayOfWeekMask))
    if kind in (OL_RECURS_MONTHLY, OL_RECURS_YEARLY):
        out["day_of_month"] = int(pattern.DayOfMonth)
    if kind in (OL_RECURS_YEARLY, OL_RECURS_YEAR_NTH):
        out["month"] = int(pattern.MonthOfYear)
    out["pattern_start"] = _iso_date(pattern.PatternStartDate)
    if bool(getattr(pattern, "NoEndDate", False)):
        out["no_end"] = True
    else:
        out["pattern_end"] = _iso_date(pattern.PatternEndDate)
        try:
            out["occurrences"] = int(pattern.Occurrences)
        except Exception:
            pass
    return out


def _iso_date(value) -> str | None:
    try:
        return value.date().isoformat() if hasattr(value, "date") else str(value)[:10]
    except Exception:
        return None


def summary(spec_or_desc: dict) -> str:
    """One line for a result: "weekly on tue, thu until 2026-12-31"."""
    kind = spec_or_desc.get("type")
    name = TYPE_NAMES.get(kind, kind) if isinstance(kind, int) else str(kind)
    interval = int(spec_or_desc.get("interval") or 1)
    parts = [name if interval == 1 else f"every {interval} {name.replace('ly', 's') if name != 'daily' else 'days'}"]
    days = spec_or_desc.get("days") or mask_days(spec_or_desc.get("day_mask") or 0)
    if days:
        parts.append("on " + ", ".join(days))
    if spec_or_desc.get("until"):
        parts.append(f"until {spec_or_desc['until']}")
    elif spec_or_desc.get("pattern_end"):
        parts.append(f"until {spec_or_desc['pattern_end']}")
    elif spec_or_desc.get("count") or spec_or_desc.get("occurrences"):
        parts.append(f"{spec_or_desc.get('count') or spec_or_desc.get('occurrences')} times")
    return " ".join(parts)
