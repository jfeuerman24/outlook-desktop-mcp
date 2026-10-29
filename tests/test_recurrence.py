"""Unit tests for tools.recurrence with a fake RecurrencePattern. No Outlook needed."""
import unittest
from datetime import date, datetime

from outlook_desktop_mcp.tools import recurrence as r

TUE = datetime(2026, 11, 3, 9, 0)   # a Tuesday
TUE_END = datetime(2026, 11, 3, 9, 30)


class _Pattern:
    """Records attribute writes in order, like the COM object would receive them."""
    def __init__(self):
        object.__setattr__(self, "writes", [])

    def __setattr__(self, name, value):
        self.writes.append((name, value))
        object.__setattr__(self, name, value)


class ParseTests(unittest.TestCase):
    def test_empty_is_none(self):
        self.assertIsNone(r.parse("", TUE))
        self.assertIsNone(r.parse("   ", TUE))

    def test_weekly_defaults_to_the_start_weekday(self):
        spec = r.parse("weekly", TUE)
        self.assertEqual(spec["type"], r.OL_RECURS_WEEKLY)
        self.assertEqual(spec["day_mask"], r.DAY_MASKS["tue"])
        self.assertEqual(spec["interval"], 1)
        self.assertIsNone(spec["until"])
        self.assertIsNone(spec["count"])

    def test_weekdays_shorthand(self):
        spec = r.parse("Weekdays", TUE)
        self.assertEqual(spec["type"], r.OL_RECURS_WEEKLY)
        self.assertEqual(r.mask_days(spec["day_mask"]), ["mon", "tue", "wed", "thu", "fri"])

    def test_monthly_and_yearly_take_the_start_date(self):
        self.assertEqual(r.parse("monthly", TUE)["day_of_month"], 3)
        y = r.parse("yearly", TUE)
        self.assertEqual((y["month_of_year"], y["day_of_month"]), (11, 3))

    def test_json_with_days_interval_and_until(self):
        spec = r.parse('{"type": "weekly", "interval": 2, "days": ["tue", "Thursday"], "until": "2026-12-31"}', TUE)
        self.assertEqual(spec["interval"], 2)
        self.assertEqual(r.mask_days(spec["day_mask"]), ["tue", "thu"])
        self.assertEqual(spec["until"], date(2026, 12, 31))

    def test_json_count_and_days_as_string(self):
        spec = r.parse('{"type": "weekly", "days": "mon, wed", "count": 6}', TUE)
        self.assertEqual(spec["count"], 6)
        self.assertEqual(r.mask_days(spec["day_mask"]), ["mon", "wed"])

    def test_errors_are_specific(self):
        with self.assertRaises(ValueError):
            r.parse("fortnightly", TUE)
        with self.assertRaises(ValueError):
            r.parse('{"type": "weekly", "days": ["funday"]}', TUE)
        with self.assertRaises(ValueError):
            r.parse('{"type": "daily", "until": "2026-12-31", "count": 3}', TUE)
        with self.assertRaises(ValueError):
            r.parse("[1, 2]", TUE)


class ApplyDescribeTests(unittest.TestCase):
    def test_apply_writes_type_before_days_and_sets_no_end(self):
        p = _Pattern()
        r.apply(p, r.parse("weekly", TUE), TUE, TUE_END)
        names = [w[0] for w in p.writes]
        self.assertLess(names.index("RecurrenceType"), names.index("DayOfWeekMask"))
        self.assertLess(names.index("DayOfWeekMask"), names.index("Interval"))
        self.assertEqual(p.PatternStartDate, "2026-11-03")
        self.assertEqual(p.StartTime, "2026-11-03 09:00:00")
        self.assertEqual(p.EndTime, "2026-11-03 09:30:00")
        self.assertTrue(p.NoEndDate)
        self.assertNotIn("Occurrences", names)

    def test_apply_count_and_until(self):
        p = _Pattern()
        r.apply(p, r.parse('{"type": "daily", "count": 5}', TUE), TUE, TUE_END)
        self.assertEqual(p.Occurrences, 5)
        p = _Pattern()
        r.apply(p, r.parse('{"type": "monthly", "until": "2027-03-01"}', TUE), TUE, TUE_END)
        self.assertEqual(p.PatternEndDate, "2027-03-01")
        self.assertEqual(p.DayOfMonth, 3)

    def test_yearly_leaves_interval_alone(self):
        p = _Pattern()
        r.apply(p, r.parse("yearly", TUE), TUE, TUE_END)
        self.assertNotIn("Interval", [w[0] for w in p.writes])
        self.assertEqual((p.MonthOfYear, p.DayOfMonth), (11, 3))

    def test_describe_and_summary(self):
        p = _Pattern()
        r.apply(p, r.parse('{"type": "weekly", "days": ["tue", "thu"], "until": "2026-12-31"}', TUE), TUE, TUE_END)
        object.__setattr__(p, "Occurrences", 17)
        d = r.describe(p)
        self.assertEqual(d["type"], "weekly")
        self.assertEqual(d["days"], ["tue", "thu"])
        self.assertEqual(d["pattern_start"], "2026-11-03")
        self.assertEqual(d["pattern_end"], "2026-12-31")
        self.assertEqual(d["occurrences"], 17)
        self.assertEqual(r.summary(d), "weekly on tue, thu until 2026-12-31")
        self.assertEqual(r.summary(r.parse('{"type": "weekly", "interval": 2, "count": 4}', TUE)),
                         "every 2 weeks on tue 4 times")


if __name__ == "__main__":
    unittest.main()
