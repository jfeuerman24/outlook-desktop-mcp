"""Unit tests for utils.html — outgoing font styling and inline answers. No Outlook needed."""
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
            self.assertEqual(h.font_style("#C00000"), "font-family:Aptos,Calibri,sans-serif;font-size:12pt;color:#C00000")

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


# A reply as Outlook's Word renderer writes the quoted message: MsoNormal paragraphs,
# <o:p> markers, entities, a list and a plain-text block with <br> line breaks.
QUOTED = (
    '<html><head><style>p.MsoNormal{margin:0}</style></head><body lang=EN-US>'
    '<div id="outlook-mcp-note"><p>See below.</p></div>'
    '<div><p class=MsoNormal>Chris, Jake &#8211; attaching the draft.<o:p></o:p></p>'
    '<ol><li class=MsoNormal><u>Term</u>: One year, consistent with LOI. <i>Confirm ok.</i></li>'
    '<li class=MsoNormal>Title: General Manager. Confirm ok.</li></ol>'
    '<p class=MsoNormal>Mike&#8217;s base is <span lang=EN-US>below</span>&nbsp;the threshold.<o:p></o:p></p>'
    '<p class=MsoPlainText>Line one<br>Line two<br>Line three</p>'
    '<p class=MsoNormal>Confirm ok.<o:p></o:p></p>'
    '</div></body></html>'
)


class InlineAnswerTests(unittest.TestCase):
    def setUp(self):
        self._env = mock.patch.dict(os.environ, {"OUTLOOK_MCP_SETTINGS_FILE": os.devnull}, clear=False)
        self._env.start()

    def tearDown(self):
        self._env.stop()

    def test_anchor_matches_across_tags_entities_and_curly_quotes(self):
        pos = h.find_anchor(QUOTED, "Mike's base is below the threshold")
        self.assertIsNotNone(pos)
        self.assertEqual(QUOTED[pos - len("threshold"):pos], "threshold")
        self.assertIsNotNone(h.find_anchor(QUOTED, "chris, jake - attaching"))
        self.assertIsNone(h.find_anchor(QUOTED, "not in the message"))
        self.assertIsNone(h.find_anchor(QUOTED, "margin:0"))  # <style> is not visible text

    def test_answer_goes_after_the_paragraph(self):
        out, missing = h.insert_inline(QUOTED, [("attaching the draft", "Thanks Mitch.")], "#C00000")
        self.assertEqual(missing, [])
        after = out.index("attaching the draft.<o:p></o:p></p>") + len("attaching the draft.<o:p></o:p></p>")
        self.assertTrue(out[after:].startswith(f'<p class="{h.INLINE_CLASS}" style="margin:0;line-height:100%;'))
        self.assertIn("color:#C00000", out[after:after + 200])
        self.assertIn("Thanks Mitch.</span></p>", out)
        self.assertLess(out.index("Thanks Mitch."), out.index("<ol>"))

    def test_answer_inside_a_list_item(self):
        out, missing = h.insert_inline(QUOTED, [("Term: One year", "Confirmed, one year.")])
        self.assertEqual(missing, [])
        li_end = out.index("</li>")
        self.assertTrue(out[:li_end].endswith("Confirmed, one year.</span></p>"))
        self.assertIn("Confirm ok.</i>", out[:li_end])

    def test_answer_after_a_line_break_is_a_line(self):
        out, missing = h.insert_inline(QUOTED, [("Line one", "Yes.")])
        self.assertEqual(missing, [])
        self.assertIn(f'Line one<br><span class="{h.INLINE_CLASS}" style="', out)
        self.assertIn(">Yes.</span><br>Line two", out)

    def test_answers_resolve_in_document_order(self):
        out, missing = h.insert_inline(QUOTED, [("Confirm ok", "First."), ("Confirm ok", "Second."), ("Confirm ok", "Third.")])
        self.assertEqual(missing, [])
        self.assertLess(out.index("First."), out.index("Title: General Manager"))
        self.assertLess(out.index("Title: General Manager"), out.index("Second."))
        self.assertLess(out.index("Line three"), out.index("Third."))

    def test_missing_anchor_changes_nothing(self):
        out, missing = h.insert_inline(QUOTED, [("Term: One year", "ok"), ("no such phrase", "x")])
        self.assertEqual(missing, ["no such phrase"])
        self.assertEqual(out, QUOTED)

    def test_multi_paragraph_answer_and_strip(self):
        out, _ = h.insert_inline(QUOTED, [("Term: One year", "Confirmed.\n\n1. one\n2. two"), ("Line one", "Yes.")])
        self.assertIn(f'<ol class="{h.INLINE_CLASS}"', out)
        self.assertEqual(h.strip_inline(out), QUOTED)
        # Word may drop the quotes around the class when the draft is reopened.
        rewritten = out.replace(f'class="{h.INLINE_CLASS}"', f'class={h.INLINE_CLASS}')
        self.assertEqual(h.strip_inline(rewritten), QUOTED)

    def test_no_block_end_appends_before_body_close(self):
        out, missing = h.insert_inline("<body><span>only text</span></body>", [("only text", "ans")])
        self.assertEqual(missing, [])
        self.assertTrue(out.endswith("ans</span></p></body>"))


if __name__ == "__main__":
    unittest.main()
