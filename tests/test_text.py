"""Unit tests for utils.text — no Outlook needed. Run: python -m unittest discover -s tests"""
import unittest

from outlook_desktop_mcp.utils.text import clean_body, strip_external_banner, trim_quoted_history

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
