"""Unit tests for utils.dasl — Restrict filter builders. No Outlook needed."""
import unittest
from datetime import datetime

from outlook_desktop_mcp.utils import dasl


class EscapeTests(unittest.TestCase):
    def test_wildcards_and_brackets_become_literals(self):
        self.assertEqual(dasl.escape_like("100%_[x]"), "100[%][_][[]x]")

    def test_quotes_doubled(self):
        self.assertEqual(dasl.escape_like("O'Brien \"q\""), "O''Brien \"\"q\"\"")

    def test_jet_date(self):
        self.assertEqual(dasl.jet_date(datetime(2026, 3, 9, 7, 5)), "03/09/2026 07:05")


class FilterTests(unittest.TestCase):
    def test_sender_filter_covers_name_directory_and_smtp(self):
        f = dasl.sender_filter("pennspring")
        for prop in (dasl.FROM_NAME, dasl.FROM_EMAIL, dasl.SENDER_SMTP):
            self.assertIn(f"\"{prop}\" LIKE '%pennspring%'", f)
        self.assertTrue(f.startswith("(") and f.endswith(")"))
        self.assertEqual(f.count(" OR "), 2)

    def test_recipient_filter_dedupes_names_and_skips_short(self):
        f = dasl.recipient_filter(["lou@pennspring.com", "Lou Castelli", "lou castelli", "x", ""])
        self.assertEqual(f.count("'%lou@pennspring.com%'"), 3)  # to, cc, bcc
        self.assertEqual(f.count("'%Lou Castelli%'"), 3)
        self.assertNotIn("'%lou castelli%'", f)
        self.assertNotIn("'%x%'", f)
        for prop in dasl.RECIPIENT_PROPS:
            self.assertIn(prop, f)

    def test_recipient_filter_accepts_single_string(self):
        self.assertIn("'%dan%'", dasl.recipient_filter("dan"))

    def test_empty_needles_rejected(self):
        with self.assertRaises(ValueError):
            dasl.recipient_filter(["", " "])

    def test_sql_joins_with_and_and_prefix(self):
        self.assertEqual(dasl.sql("a", "", "b"), "@SQL=a AND b")
        self.assertEqual(dasl.sql("", ""), "")

    def test_subject_and_body(self):
        self.assertEqual(dasl.subject_filter("EHC"), f"\"{dasl.SUBJECT}\" LIKE '%EHC%'")
        self.assertEqual(dasl.body_filter("EHC"), f"\"{dasl.BODY}\" LIKE '%EHC%'")
        self.assertEqual(dasl.body_filter("O'Neil", indexed=True), f"\"{dasl.BODY}\" ci_phrasematch 'O''Neil'")
        both = dasl.subject_or_body_filter("x", indexed_body=True)
        self.assertIn("LIKE '%x%'", both)
        self.assertIn("ci_phrasematch 'x'", both)

    def test_date_bounds(self):
        d = datetime(2026, 10, 1, 0, 0)
        self.assertEqual(dasl.received_since(d), f"\"{dasl.DATE_RECEIVED}\" >= '10/01/2026 00:00'")
        self.assertEqual(dasl.received_until(d), f"\"{dasl.DATE_RECEIVED}\" <= '10/01/2026 00:00'")

    def test_looks_like_address(self):
        self.assertTrue(dasl.looks_like_address("lou@pennspring.com"))
        self.assertTrue(dasl.looks_like_address("pennspring.com"))
        self.assertFalse(dasl.looks_like_address("Lou Castelli"))
        self.assertFalse(dasl.looks_like_address("dan"))
        self.assertFalse(dasl.looks_like_address("Sanders & Johnson Inc."))


if __name__ == "__main__":
    unittest.main()
