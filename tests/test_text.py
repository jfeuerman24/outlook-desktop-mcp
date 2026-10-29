"""Unit tests for utils.text — no Outlook needed. Run: python -m unittest discover -s tests"""
import unittest

from outlook_desktop_mcp.utils.text import clean_body, shorten_links, strip_external_banner, trim_quoted_history

OUTLOOK_REPLY = (
    "Sounds good, let's do Thursday.\r\n\r\nLinda\r\n\r\n"
    "From: Alex Morgan <alex@company.example.com>\r\nSent: Monday, October 26, 2026 1:48 PM\r\n"
    "To: Linda Harlow <linda@example.com>\r\nSubject: RE: Requested Info\r\n\r\nLinda,\r\n\r\nSee attached.\r\n"
)
GMAIL_REPLY = (
    "Works for me.\n\nOn Mon, Oct 26, 2026 at 5:40 PM Linda Harlow <linda@example.com>\nwrote:\n\n> Hi Alex;\n> Thanks\n"
)
ORIGINAL_MESSAGE = "Yes.\n\n-----Original Message-----\nFrom: Bob\nSent: Monday\n\nQuestion?"
SEPARATOR = "Approved.\n\n________________________________\nFrom: Bob\nSent: Monday\n"
BARE_FORWARD = "\n\nFrom: Bob <bob@example.com>\nSent: Monday\nTo: Alex\nSubject: FW: deck\n\nHere is the deck.\n"
BANNER = (
    "You don't often get email from linda@example.com. Learn why this is important\n"
    "Hi Alex;\nThanks for the info.\n"
)


TRACKING = "https://click.jetblue.com/f/a/" + "x" * 120 + "?utm_source=email&token=" + "y" * 80
IMAGE = "https://braze-images.com/appboy/assets/" + "z" * 60 + "/original.png?1761700000"
CONFIRMATION = (
    "Your Flight Itinerary\nYour JetBlue confirmation code is JUVTNS\n\n"
    f" <{IMAGE}>\nBOS  <{IMAGE}>\n  RDU\nFlight 2283\nTue, Nov 3 11:36am\n\n"
    f"Manage your booking <{TRACKING}>\nCheck in <{TRACKING}>\n"
    f"Questions? https://jetblue.com/help or call 1-800-JETBLUE.\n"
    f"Unsubscribe <{TRACKING}> Privacy <{TRACKING}>\n"
)


class LinkTests(unittest.TestCase):
    def test_long_links_become_domain_placeholders_and_short_ones_stay(self):
        out, count = shorten_links(CONFIRMATION)
        self.assertEqual(count, 6)
        self.assertIn("Manage your booking [link: jetblue.com]", out)
        self.assertIn("https://jetblue.com/help", out)
        self.assertNotIn("utm_source", out)
        self.assertNotIn("x" * 20, out)

    def test_image_links_are_dropped_and_repeats_collapsed(self):
        out, _ = shorten_links(CONFIRMATION)
        self.assertNotIn("braze-images", out)
        self.assertIn("Flight 2283", out)
        self.assertEqual(out.count("[link: jetblue.com]"), 4)  # a word between two links keeps both
        adjacent, _ = shorten_links(f"footer <{TRACKING}>\n <{TRACKING}>  <{TRACKING}> end")
        self.assertEqual(adjacent, "footer [link: jetblue.com] end")
        self.assertEqual(shorten_links("no links here"), ("no links here", 0))

    def test_clean_body_shortens_before_the_cap_and_reports_totals(self):
        out = clean_body(CONFIRMATION)
        self.assertFalse(out["body_truncated"])
        self.assertEqual(out["links_shortened"], 6)
        self.assertLess(len(out["body"]), 400)
        self.assertNotIn("body_total_chars", out)
        small = clean_body(CONFIRMATION, max_length=100)
        self.assertTrue(small["body_truncated"])
        self.assertEqual(small["body_total_chars"], len(out["body"]))
        self.assertTrue(small["body"].endswith("... [truncated]"))
        self.assertFalse(clean_body("x" * 9000, max_length=0)["body_truncated"])
        self.assertTrue(clean_body("x" * 9000)["body_truncated"])  # default cap 8,000


class TrimTests(unittest.TestCase):
    def test_outlook_header_block(self):
        body, trimmed = trim_quoted_history(OUTLOOK_REPLY)
        self.assertTrue(trimmed)
        self.assertEqual(body, "Sounds good, let's do Thursday.\n\nLinda")

    def test_gmail_wrote_line_even_when_wrapped(self):
        body, trimmed = trim_quoted_history(GMAIL_REPLY)
        self.assertTrue(trimmed)
        self.assertEqual(body, "Works for me.")

    def test_original_message_and_separator(self):
        self.assertEqual(trim_quoted_history(ORIGINAL_MESSAGE), ("Yes.", True))
        self.assertEqual(trim_quoted_history(SEPARATOR), ("Approved.", True))

    def test_bare_forward_keeps_everything(self):
        body, trimmed = trim_quoted_history(BARE_FORWARD)
        self.assertFalse(trimmed)
        self.assertIn("Here is the deck.", body)
        self.assertTrue(body.startswith("From: Bob"))

    def test_plain_message_untouched(self):
        body, trimmed = trim_quoted_history("Hi,\n\nFrom my side all good.\n\nThanks")
        self.assertFalse(trimmed)  # "From my side" is not a header block
        self.assertEqual(body, "Hi,\n\nFrom my side all good.\n\nThanks")

    def test_banner_removed(self):
        self.assertEqual(strip_external_banner(BANNER).strip(), "Hi Alex;\nThanks for the info.")

    def test_clean_body_flags_and_cap(self):
        record = clean_body(OUTLOOK_REPLY, max_length=10)
        self.assertEqual(record["body_trimmed"], True)
        self.assertEqual(record["body_truncated"], True)
        self.assertTrue(record["body"].endswith("... [truncated]"))
        raw = clean_body(OUTLOOK_REPLY, trim_history=False)
        self.assertFalse(raw["body_trimmed"])
        self.assertIn("See attached.", raw["body"])


if __name__ == "__main__":
    unittest.main()
