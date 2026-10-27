"""Every calendar folder in a store, not just the default one.

Outlook keeps extra calendars (Birthdays, holidays, calendars shared with you, internet
subscriptions) as folders whose ``DefaultItemType`` is appointment, usually nested under
the default Calendar folder. ``list_calendar_folders`` walks the store's folder tree for
them; ``resolve_calendars`` turns a tool's ``calendar`` argument into concrete folders.

The walk touches every folder in the store. On a mailbox that is not cached locally that
is a server round trip per folder and was measured at 18 to 30 seconds, so the result is
cached per store for ``CACHE_TTL`` seconds. All callers run on the single COM thread, so
holding the folder objects is safe; a folder that has since vanished raises when used,
and the caller then calls ``invalidate``.
"""
import time

from outlook_desktop_mcp.tools._folder_constants import OL_FOLDER_CALENDAR

OL_ITEM_TYPE_APPOINTMENT = 1
MAX_DEPTH = 8
MAX_FOLDERS = 300
CACHE_TTL = 600.0  # seconds
_SKIP_FOLDERS = {"deleted items", "junk email", "sync issues", "conversation history"}
ALL_CALENDARS = {"all", "*"}
_PEOPLE_FREE = ("birthday", "holiday")

_cache: dict[str, tuple[float, list]] = {}


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


def _store_key(store) -> str:
    try:
        return str(store.StoreID)
    except Exception:
        return str(id(store))


def invalidate(store=None) -> None:
    """Forget the cached calendar list for ``store`` (or every store)."""
    if store is None:
        _cache.clear()
    else:
        _cache.pop(_store_key(store), None)


def list_calendar_folders(store, refresh: bool = False) -> list[tuple[str, object, bool]]:
    """``(path, folder, is_default)`` for every calendar folder in the store, default first.

    Served from the per-store cache unless it is older than ``CACHE_TTL`` or ``refresh``.
    """
    key = _store_key(store)
    now = time.monotonic()
    if not refresh:
        hit = _cache.get(key)
        if hit and now - hit[0] < CACHE_TTL:
            return hit[1]
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
    _cache[key] = (now, calendars)
    return calendars


def describe_calendars(store) -> list[dict]:
    """Live listing for the user (bypasses the cache and refreshes it)."""
    out = []
    for path, folder, is_default in list_calendar_folders(store, refresh=True):
        entry = {"path": path, "name": path.rpartition("/")[2], "is_default": is_default}
        try:
            entry["item_count"] = folder.Items.Count
        except Exception:
            pass
        out.append(entry)
    return out


def is_people_calendar(path: str) -> bool:
    """False for Birthdays and holiday calendars, which never hold meetings with people."""
    low = (path or "").lower()
    return not any(word in low for word in _PEOPLE_FREE)


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
