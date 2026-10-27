"""Unit tests for utils.deadline and the bridge's deadline arithmetic. No Outlook needed."""
import threading
import time
import unittest

from outlook_desktop_mcp.utils import deadline


class DeadlineTests(unittest.TestCase):
    def tearDown(self):
        deadline.set_current(None)

    def test_never_expires_when_unset(self):
        self.assertFalse(deadline.Deadline(None).expired)
        self.assertIsNone(deadline.Deadline(None).remaining())
        self.assertFalse(deadline.expired())  # nothing installed on this thread

    def test_expires_after_seconds(self):
        d = deadline.Deadline(0.05)
        self.assertFalse(d.expired)
        self.assertGreater(d.remaining(), 0)
        time.sleep(0.08)
        self.assertTrue(d.expired)
        self.assertEqual(d.remaining(), 0.0)

    def test_absolute_point(self):
        d = deadline.Deadline(at=time.monotonic() - 1)
        self.assertTrue(d.expired)

    def test_current_is_per_thread(self):
        deadline.set_current(deadline.Deadline(at=time.monotonic() - 1))
        self.assertTrue(deadline.expired())
        seen = {}

        def other():
            seen["expired"] = deadline.expired()
        t = threading.Thread(target=other)
        t.start()
        t.join()
        self.assertFalse(seen["expired"])

    def test_partial_note_mentions_narrowing(self):
        note = deadline.partial_note("The sweep")
        self.assertTrue(note.startswith("The sweep stopped early"))
        self.assertIn("call_timeout_s", note)


class BridgeRequestTests(unittest.TestCase):
    """The request bookkeeping is pure Python; the COM thread itself needs Outlook."""

    def test_soft_deadline_sits_before_hard(self):
        from outlook_desktop_mcp.com_bridge import _Request, SOFT_MARGIN
        r = _Request(lambda *a: None, (), {}, 60.0)
        self.assertAlmostEqual(r.deadline_at - r.soft_at, SOFT_MARGIN, places=2)
        short = _Request(lambda *a: None, (), {}, 4.0)
        self.assertAlmostEqual(short.deadline_at - short.soft_at, 1.0, places=2)  # 25% of a short timeout
        self.assertIsNone(short.started)

    def test_default_timeout_clamped_and_from_setting(self):
        import os
        from unittest import mock
        from outlook_desktop_mcp import com_bridge
        with mock.patch.dict(os.environ, {"OUTLOOK_MCP_CALL_TIMEOUT_S": "90"}):
            self.assertEqual(com_bridge.default_timeout(), 90.0)
        with mock.patch.dict(os.environ, {"OUTLOOK_MCP_CALL_TIMEOUT_S": "2"}):
            self.assertEqual(com_bridge.default_timeout(), com_bridge.MIN_TIMEOUT)
        with mock.patch.dict(os.environ, {"OUTLOOK_MCP_CALL_TIMEOUT_S": "banana"}):
            self.assertEqual(com_bridge.default_timeout(), com_bridge.DEFAULT_TIMEOUT)

    def test_busy_for_reports_overrun_only(self):
        from outlook_desktop_mcp.com_bridge import OutlookBridge, _Request
        b = OutlookBridge()
        self.assertIsNone(b.busy_for())
        r = _Request(lambda *a: None, (), {}, 60.0)
        r.started = time.monotonic()
        b._running = r
        self.assertIsNone(b.busy_for())  # within its limit
        r.deadline_at = time.monotonic() - 3
        self.assertGreaterEqual(b.busy_for(), 3)


if __name__ == "__main__":
    unittest.main()
