"""Unit tests for utils.html — outgoing font styling. No Outlook needed."""
import os
import unittest
from unittest import mock

from outlook_desktop_mcp.utils import html as h


class StyleTests(unittest.TestCase):
    def setUp(self):
        # Never read the machine's real settings file during tests.
        self._env = mock.patch.dict(os.environ, {"OUTLOOK_MCP_SETTINGS_FILE": os.devnull}, clear=False)
        self._env.start()

    def tearDown(self):
        self._env.stop()

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
        self.assertEqual(out.count("<p "), 5)  # 3 paragraphs + 2 blank-line spacers
        self.assertEqual(out.count("&nbsp;"), 2)
        self.assertIn("Thanks &amp; see &lt;attached&gt;.<br>Talk Thursday.", out)
        self.assertIn('font-family:Aptos', out)
        self.assertTrue(out.startswith('<p style="margin:0;line-height:100%;font-family:Aptos'))

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

    def test_text_to_html_lists_quotes_and_bold(self):
        out = h.text_to_html(
            "Steps:\n\n1. First **thing**\n2. Second\n   continued\n\n- a\n- b\n\n> paste me\n> and me\n\nDone & dusted."
        )
        self.assertIn("<ol", out)
        self.assertIn("<li", out)
        self.assertIn("First <b>thing</b>", out)
        self.assertIn("Second<br>continued", out)
        self.assertIn("<ul", out)
        self.assertEqual(out.count("<li"), 4)
        self.assertIn("margin-left:24pt", out)
        self.assertIn("line-height:100%", out)
        self.assertEqual(out.count("&nbsp;"), 4)  # one blank line between each of the 5 blocks
        self.assertIn("paste me<br>and me", out)
        self.assertIn("Done &amp; dusted.", out)
        self.assertNotIn("**", out)

    def test_note_marker_round_trip(self):
        doc = ('<html><body>' + h.wrap_note('<p>old note</p><div><p>nested</p></div>')
               + '<div class="sig">-- Jake</div><div>quoted history</div></body></html>')
        span = h.find_note(doc)
        self.assertIsNotNone(span)
        start, end = span
        self.assertTrue(doc[start:end].startswith('<div id="outlook-mcp-note">'))
        self.assertTrue(doc[start:end].endswith('</div>'))
        self.assertIn('nested', doc[start:end])
        self.assertNotIn('-- Jake', doc[start:end])
        new = h.replace_note(doc, '<p>new note</p>')
        self.assertIn('<p>new note</p>', new)
        self.assertNotIn('old note', new)
        self.assertIn('-- Jake', new)
        self.assertIn('quoted history', new)

    def test_note_marker_survives_word_style_rewrite(self):
        # Outlook's editor may drop the quotes and add attributes to the div.
        doc = '<body><div id=outlook-mcp-note class=WordSection1><p>x</p></div><p>sig</p></body>'
        self.assertEqual(h.replace_note(doc, '<p>y</p>'), '<body><div id="outlook-mcp-note"><p>y</p></div><p>sig</p></body>')

    def test_replace_note_without_marker_returns_none(self):
        self.assertIsNone(h.replace_note('<body><p>hand written</p></body>', '<p>y</p>'))

