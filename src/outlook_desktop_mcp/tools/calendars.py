"""Every calendar folder in a store, not just the default one.

Outlook keeps extra calendars (Birthdays, holidays, calendars shared with you, internet
subscriptions) as folders whose ``DefaultItemType`` is appointment, usually nested under
the default Calendar folder. ``list_calendar_folders`` walks the store's folder tree for
them; ``resolve_calendars`` turns a tool's ``calendar`` argument into concrete folders.
"""
from outlook_desktop_mcp.tools._folder_constants import OL_FOLDER_CALENDAR

OL_ITEM_TYPE_APPOINTMENT = 1
MAX_DEPTH = 8
MAX_FOLDERS = 300
_SKIP_FOLDERS = {"deleted items", "junk email", "sync issues", "conversation history"}
ALL_CALENDARS = {"all", "*"}


def _walk(folder, path: str, out: list, depth: int = 0) -> None:
    if depth > MAX_DEPTH or len(out) >= MAX_FOLDERS:
        return
    try:
        count = folder.Folders.Count
    except Exception:
        return
    for i in range(count):
        try:
            sub = folder.Folders.Item(i + 1)
            name = str(sub.Name or "").strip()
        except Exception:
            continue
        if not name or name.lower() in _SKIP_FOLDERS:
            continue
        sub_path = f"{path}/{name}" if path else name
        try:
            if int(sub.DefaultItemType) == OL_ITEM_TYPE_APPOINTMENT:
                out.append((sub_path, sub))
        except Exception:
            pass
        _walk(sub, sub_path, out, depth + 1)


def list_calendar_folders(store) -> list[tuple[str, object, bool]]:
    """``(path, folder, is_default)`` for every calendar folder in the store, default first."""
    default_id = ""
    try:
        default_id = str(store.GetDefaultFolder(OL_FOLDER_CALENDAR).EntryID)
    except Exception:
        pass
    found: list[tuple[str, object]] = []
    try:
        _walk(store.GetRootFolder(), "", found)
    except Exception:
        pass
    calendars = []
    for path, folder in found:
        try:
            is_default = bool(default_id) and str(folder.EntryID) == default_id
        except Exception:
            is_default = False
        calendars.append((path, folder, is_default))
    calendars.sort(key=lambda entry: not entry[2])  # stable: default first, then tree order
    return calendars


def describe_calendars(store) -> list[dict]:
    out = []
    for path, folder, is_default in list_calendar_folders(store):
        entry = {"path": path, "name": path.rpartition("/")[2], "is_default": is_default}
        try:
            entry["item_count"] = folder.Items.Count
        except Exception:
            pass
        out.append(entry)
    return out


def resolve_calendars(store, calendar: str = "") -> list[tuple[str, object]]:
    """Folders for a tool's ``calendar`` argument.

    "" → the default Calendar; "all" or "*" → every calendar in the store; otherwise a
    case-insensitive match on the calendar's path (``Calendar/Family``) or name
    (``Family``). Raises ``ValueError`` when nothing matches.
    """
    wanted = (calendar or "").strip()
    if not wanted:
        return [("Calendar", store.GetDefaultFolder(OL_FOLDER_CALENDAR))]
    calendars = list_calendar_folders(store)
    if wanted.lower() in ALL_CALENDARS:
        return [(path, folder) for path, folder, _default in calendars]
    wanted_lower = wanted.lower().strip("/")
    for path, folder, _default in calendars:
        if path.lower() == wanted_lower:
            return [(path, folder)]
    for path, folder, _default in calendars:
        if path.rpartition("/")[2].lower() == wanted_lower:
            return [(path, folder)]
    for path, folder, _default in calendars:
        if wanted_lower in path.lower():
            return [(path, folder)]
    available = ", ".join(path for path, _f, _d in calendars) or "(none)"
    raise ValueError(f"Calendar '{calendar}' not found. Available: {available}. Use list_calendars.")
