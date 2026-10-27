"""Unit tests for tools.calendars with a fake store: cache, refresh, people-free calendars."""
import unittest

from outlook_desktop_mcp.tools import calendars


class _Folders:
    def __init__(self, items):
        self._items = items

    @property
    def Count(self):
        return len(self._items)

    def Item(self, i):
        return self._items[i - 1]


class _Folder:
    def __init__(self, name, item_type=0, children=(), entry_id=None):
        self.Name = name
        self.DefaultItemType = item_type
        self.EntryID = entry_id or f"id-{name}"
        self.Folders = _Folders(list(children))
        self.Items = _Folders([])


class _Store:
    walks = 0

    def __init__(self, store_id, root, default_cal):
        self.StoreID = store_id
        self._root = root
        self._default = default_cal

    def GetRootFolder(self):
        _Store.walks += 1
        return self._root

    def GetDefaultFolder(self, _kind):
        return self._default


def _make_store(store_id="S1"):
    cal = _Folder("Calendar", 1, children=[_Folder("Birthdays", 1), _Folder("Family", 1)])
    root = _Folder("root", children=[_Folder("Inbox", 0), cal, _Folder("United States holidays", 1)])
    return _Store(store_id, root, cal)


class CalendarTests(unittest.TestCase):
    def setUp(self):
        calendars.invalidate()
        _Store.walks = 0

    def test_lists_every_calendar_default_first(self):
        paths = [p for p, _f, _d in calendars.list_calendar_folders(_make_store())]
        self.assertEqual(paths, ["Calendar", "Calendar/Birthdays", "Calendar/Family", "United States holidays"])
        self.assertTrue(calendars.list_calendar_folders(_make_store())[0][2])

    def test_cache_avoids_repeated_walks(self):
        store = _make_store()
        calendars.list_calendar_folders(store)
        calendars.list_calendar_folders(store)
        calendars.resolve_calendars(store, "all")
        self.assertEqual(_Store.walks, 1)
        calendars.list_calendar_folders(store, refresh=True)
        self.assertEqual(_Store.walks, 2)
        calendars.invalidate(store)
        calendars.list_calendar_folders(store)
        self.assertEqual(_Store.walks, 3)

    def test_cache_is_per_store(self):
        calendars.list_calendar_folders(_make_store("A"))
        calendars.list_calendar_folders(_make_store("B"))
        self.assertEqual(_Store.walks, 2)

    def test_describe_refreshes(self):
        store = _make_store()
        calendars.list_calendar_folders(store)
        calendars.describe_calendars(store)
        self.assertEqual(_Store.walks, 2)

    def test_resolve_by_name_path_and_default(self):
        store = _make_store()
        self.assertEqual(calendars.resolve_calendars(store, "")[0][0], "Calendar")
        self.assertEqual(calendars.resolve_calendars(store, "family")[0][0], "Calendar/Family")
        self.assertEqual(calendars.resolve_calendars(store, "Calendar/Birthdays")[0][0], "Calendar/Birthdays")
        with self.assertRaises(ValueError):
            calendars.resolve_calendars(store, "nope")

    def test_people_calendars_exclude_birthdays_and_holidays(self):
        self.assertTrue(calendars.is_people_calendar("Calendar"))
        self.assertTrue(calendars.is_people_calendar("Calendar/Kate Work"))
        self.assertFalse(calendars.is_people_calendar("Calendar/Birthdays"))
        self.assertFalse(calendars.is_people_calendar("United States holidays"))


if __name__ == "__main__":
    unittest.main()
