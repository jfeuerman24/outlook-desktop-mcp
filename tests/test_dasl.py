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


class FlagFilterTests(unittest.TestCase):
    def test_flagged_is_the_open_task_or_the_old_flag(self):
        f = dasl.flagged_filter()
        self.assertEqual(f, f'("{dasl.TASK_COMPLETE}" = 0 OR "{dasl.FLAG_STATUS}" = 2)')
        self.assertIn("811C000B", dasl.TASK_COMPLETE)  # PidLidTaskComplete, PSETID_Task

    def test_unread_and_flagged_chain_in_one_sql(self):
        s = dasl.sql(dasl.flagged_filter(), dasl.unread_filter())
        self.assertTrue(s.startswith("@SQL=("))
        self.assertIn(f' AND "{dasl.READ}" = 0', s)


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

    def test_people_filter_covers_sender_and_recipients(self):
        f = dasl.people_filter("earlgrayandsons.com", ["Mike Gray"])
        for prop in dasl.SENDER_PROPS + dasl.RECIPIENT_PROPS:
            self.assertIn(f"\"{prop}\" LIKE '%earlgrayandsons.com%'", f)
        self.assertEqual(f.count("'%Mike Gray%'"), 3)  # recipients only
        self.assertTrue(f.startswith("(") and f.endswith(")"))

    def test_anywhere_filter_subject_sender_recipients_and_optional_body(self):
        f = dasl.anywhere_filter("Earl Gray")
        self.assertIn(f"\"{dasl.SUBJECT}\" LIKE '%Earl Gray%'", f)
        self.assertIn(f"\"{dasl.SENDER_SMTP}\" LIKE '%Earl Gray%'", f)
        self.assertIn(f"\"{dasl.DISPLAY_TO}\" LIKE '%Earl Gray%'", f)
        self.assertNotIn(dasl.BODY, f)
        self.assertIn("ci_phrasematch 'Earl Gray'", dasl.anywhere_filter("Earl Gray", indexed_body=True))
        self.assertIn(f"\"{dasl.BODY}\" LIKE '%Earl Gray%'", dasl.anywhere_filter("Earl Gray", indexed_body=False))
        # A one-character needle cannot go through the recipient filter (minimum two).
        self.assertNotIn(dasl.DISPLAY_TO, dasl.anywhere_filter("x"))

    def test_looks_like_address(self):
        self.assertTrue(dasl.looks_like_address("lou@pennspring.com"))
        self.assertTrue(dasl.looks_like_address("pennspring.com"))
        self.assertFalse(dasl.looks_like_address("Lou Castelli"))
        self.assertFalse(dasl.looks_like_address("dan"))
        self.assertFalse(dasl.looks_like_address("Sanders & Johnson Inc."))


if __name__ == "__main__":
    unittest.main()
