"""Unit tests for utils.html — outgoing font styling. No Outlook needed."""
import os
import unittest
from unittest import mock

from outlook_desktop_mcp.utils import html as h


class StyleTests(unittest.TestCase):
    def test_defaults_are_aptos_12_black(self):
        with mock.patch.dict(os.environ, {}, clear=False):
            for key in ("OUTLOOK_MCP_FONT_FAMILY", "OUTLOOK_MCP_FONT_SIZE_PT", "OUTLOOK_MCP_FONT_COLOR"):
                os.environ.pop(key, None)
            self.assertEqual(h.font_style(), "font-family:Aptos,Calibri,sans-serif;font-size:12pt;color:#000000")

    def test_env_override(self):
        with mock.patch.dict(os.environ, {"OUTLOOK_MCP_FONT_FAMILY": "Calibri", "OUTLOOK_MCP_FONT_SIZE_PT": "11",
                                          "OUTLOOK_MCP_FONT_COLOR": "#333333"}):
            self.assertEqual(h.font_style(), "font-family:Calibri,Calibri,sans-serif;font-size:11pt;color:#333333")

    def test_text_to_html_paragraphs_and_escaping(self):
        out = h.text_to_html("Hi Linda,\n\nThanks & see <attached>.\nTalk Thursday.\r\n\r\nAlex")
        self.assertEqual(out.count("<p "), 3)
        self.assertIn("Thanks &amp; see &lt;attached&gt;.<br>Talk Thursday.", out)
        self.assertIn('font-family:Aptos', out)
        self.assertTrue(out.startswith('<p style="margin:0 0 12pt 0;font-family:Aptos'))

    def test_inject_into_body_sits_above_quoted_history(self):
        doc = '<html><head></head><body lang="EN-US" style="x"><div>From: old</div></body></html>'
        out = h.inject_into_body(doc, "<p>new</p>")
        self.assertLess(out.index("<p>new</p>"), out.index("<div>From: old</div>"))
        self.assertTrue(out.startswith('<html><head></head><body lang="EN-US" style="x">'))
        self.assertEqual(h.inject_into_body("", "<p>x</p>"), "<p>x</p>")

    def test_style_html_wraps(self):
        self.assertTrue(h.style_html("<b>hi</b>").startswith('<div style="font-family:Aptos'))
        self.assertIn("<b>hi</b></div>", h.style_html("<b>hi</b>"))


if __name__ == "__main__":
    unittest.main()
