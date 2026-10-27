"""Unit tests for utils.dates — no Outlook needed."""
import re
import unittest
from datetime import datetime

from outlook_desktop_mcp.utils.dates import iso, iso_date, to_local


class _ComDate:
    """Mimics a pywintypes datetime: has the date attributes, is not a datetime."""

    def __init__(self, *parts):
        self.year, self.month, self.day, self.hour, self.minute, self.second = parts


class DateTests(unittest.TestCase):
    def test_iso_has_offset(self):
        text = iso(_ComDate(2026, 10, 26, 17, 40, 5))
        self.assertRegex(text, r"^2026-10-26T17:40:05[+-]\d{2}:\d{2}$")
        self.assertEqual(iso(datetime(2026, 1, 2, 3, 4, 5))[:19], "2026-01-02T03:04:05")

    def test_no_date_sentinel(self):
        self.assertIsNone(iso(_ComDate(4501, 1, 1, 0, 0, 0)))
        self.assertIsNone(iso(None))
        self.assertIsNone(to_local("garbage"))

    def test_iso_date(self):
        self.assertEqual(iso_date(_ComDate(2026, 3, 1, 23, 59, 0)), "2026-03-01")
        self.assertIsNone(iso_date(_ComDate(4501, 1, 1, 0, 0, 0)))


if __name__ == "__main__":
    unittest.main()
