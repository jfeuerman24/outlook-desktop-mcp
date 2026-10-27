"""Unit tests for the pure helpers in tools.conversation. No Outlook needed."""
import unittest

from outlook_desktop_mcp.tools.conversation import thread_depth

ROOT = "01DC3F2A9B1C4D5E6F708192A3B4C5D6E7F8091A2B3C"  # 22-byte header as hex: the first message
LEVEL = "0123456789"                                     # 5 bytes as hex: one reply level


class ThreadDepthTests(unittest.TestCase):
    def test_first_message_is_depth_zero(self):
        self.assertEqual(len(ROOT), 44)
        self.assertEqual(thread_depth(ROOT), 0)

    def test_each_reply_adds_one_level(self):
        self.assertEqual(thread_depth(ROOT + LEVEL), 1)
        self.assertEqual(thread_depth(ROOT + LEVEL + "ABCDEF0123"), 2)
        self.assertEqual(thread_depth(ROOT.lower() + LEVEL), 1)

    def test_unknown_when_missing_or_malformed(self):
        self.assertIsNone(thread_depth(None))
        self.assertIsNone(thread_depth(""))
        self.assertIsNone(thread_depth(ROOT[:-2]))        # shorter than a header
        self.assertIsNone(thread_depth(ROOT + "01234"))   # not a whole level
        self.assertIsNone(thread_depth("Z" * 44))         # not hex


if __name__ == "__main__":
    unittest.main()
