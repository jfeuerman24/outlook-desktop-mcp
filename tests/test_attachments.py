"""Unit tests for the attachment helpers, the open-item scan and COM error wording. No Outlook needed."""
import os
import tempfile
import unittest

from outlook_desktop_mcp import server
from outlook_desktop_mcp.utils.errors import NO_DETAIL_HINT, format_com_error

try:
    import pywintypes
except ImportError:  # not Windows
    pywintypes = None


def _touch(path: str, data: bytes = b"") -> str:
    with open(path, "wb") as f:
        f.write(data)
    return path


class FakeAttachments:
    """Outlook's Attachments collection; refuses files whose name is in ``refuse_names``."""

    def __init__(self, refuse_names=()):
        self.refuse_names = set(refuse_names)
        self.added: list[str] = []

    @property
    def Count(self):
        return len(self.added)

    def Add(self, path):
        if os.path.basename(path) in self.refuse_names:
            raise RuntimeError("Outlook could not read it")
        self.added.append(path)
        return path


class FakeItem:
    def __init__(self, refuse_names=()):
        self.Attachments = FakeAttachments(refuse_names)


class ResolvePathsTests(unittest.TestCase):
    def test_missing_file_raises_before_anything_is_created(self):
        with self.assertRaises(ValueError) as ctx:
            server._resolve_attachment_paths(r"C:\nowhere\no-such-file.xlsx")
        self.assertIn("Attachment not found", str(ctx.exception))

    def test_splits_on_semicolons_and_strips_quotes(self):
        with tempfile.TemporaryDirectory() as d:
            a = _touch(os.path.join(d, "a.txt"), b"a")
            b = _touch(os.path.join(d, "b.txt"), b"b")
            self.assertEqual(server._resolve_attachment_paths(f'"{a}"; {b} ;'), [a, b])

    def test_empty_means_no_files(self):
        self.assertEqual(server._resolve_attachment_paths(""), [])
        self.assertEqual(server._resolve_attachment_paths(" ; "), [])


class OwnerFileTests(unittest.TestCase):
    def test_excel_and_word_owner_files_are_recognised(self):
        with tempfile.TemporaryDirectory() as d:
            book = _touch(os.path.join(d, "Model.xlsx"), b"x")
            self.assertFalse(server._office_owner_file(book))
            _touch(os.path.join(d, "~$Model.xlsx"))
            self.assertTrue(server._office_owner_file(book))
            doc = _touch(os.path.join(d, "Memo.docx"), b"x")
            _touch(os.path.join(d, "~$mo.docx"))  # Word drops the first two characters
            self.assertTrue(server._office_owner_file(doc))


class AttachFilesTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.book = _touch(os.path.join(self.tmp.name, "Model.xlsx"), b"12345")

    def tearDown(self):
        self.tmp.cleanup()

    def test_a_readable_file_is_attached_in_place(self):
        item = FakeItem()
        added, failed, tmpdir = server._attach_files(item, [self.book])
        self.assertEqual(failed, [])
        self.assertIsNone(tmpdir)
        self.assertEqual(item.Attachments.added, [self.book])
        self.assertEqual(added[0]["filename"], "Model.xlsx")
        self.assertEqual(added[0]["size"], 5)
        self.assertNotIn("note", added[0])

    def test_a_file_outlook_cannot_read_in_place_is_attached_as_a_copy_with_the_same_name(self):
        item = FakeItem()
        item.Attachments.Add = _refuse_once(item.Attachments, self.book)
        _touch(os.path.join(self.tmp.name, "~$Model.xlsx"))
        added, failed, tmpdir = server._attach_files(item, [self.book])
        try:
            self.assertEqual(failed, [])
            self.assertIsNotNone(tmpdir)
            copy = item.Attachments.added[0]
            self.assertNotEqual(copy, self.book)
            self.assertEqual(os.path.basename(copy), "Model.xlsx")
            self.assertTrue(copy.startswith(tmpdir))
            with open(copy, "rb") as f:
                self.assertEqual(f.read(), b"12345")
            self.assertIn("identical copy", added[0]["note"])
            self.assertIn("Excel", added[0]["note"])
        finally:
            server._rm_tmpdir(tmpdir)
        self.assertFalse(os.path.exists(tmpdir))

    def test_a_file_that_fails_twice_is_reported_and_the_others_still_attach(self):
        ok = _touch(os.path.join(self.tmp.name, "ok.txt"), b"ok")
        item = FakeItem(refuse_names={"Model.xlsx"})  # original and copy share the name, so both fail
        added, failed, tmpdir = server._attach_files(item, [self.book, ok])
        server._rm_tmpdir(tmpdir)
        self.assertEqual([a["filename"] for a in added], ["ok.txt"])
        self.assertEqual(len(failed), 1)
        self.assertEqual(failed[0]["path"], self.book)
        self.assertEqual(failed[0]["error"], "RuntimeError: Outlook could not read it")
        self.assertIn("copy could not be attached", failed[0]["reason"])

    def test_extras_warn_only_when_something_failed(self):
        self.assertEqual(server._attachment_extras([{"filename": "a"}], []), {"files_attached": [{"filename": "a"}]})
        extra = server._attachment_extras([], [{"path": "x", "error": "e", "reason": "r"}])
        self.assertIn("files_failed", extra)
        self.assertIn("add_attachment", extra["warning"])


def _refuse_once(attachments: FakeAttachments, path: str):
    """An Add that fails for ``path`` itself but accepts a copy of it elsewhere."""
    real = attachments.Add

    def add(p):
        if p == path:
            raise RuntimeError("in use")
        return real(p)
    return add


class FakeNamespace:
    def CompareEntryIDs(self, a, b):
        return a == b


class FakeCollection:
    def __init__(self, items):
        self._items = list(items)

    @property
    def Count(self):
        return len(self._items)

    def Item(self, i):
        return self._items[i - 1]


class FakeInspector:
    def __init__(self, item):
        self.CurrentItem = item


class FakeMail:
    def __init__(self, entry_id):
        self.EntryID = entry_id


class FakeExplorer:
    def __init__(self, selected):
        self.Selection = FakeCollection(selected)


class FakeOutlook:
    def __init__(self, inspectors=(), selected=None):
        self.Inspectors = FakeCollection(FakeInspector(i) for i in inspectors)
        self._explorer = FakeExplorer(selected) if selected is not None else None

    def ActiveExplorer(self):
        return self._explorer


class OpenInOutlookTests(unittest.TestCase):
    def test_nowhere(self):
        outlook = FakeOutlook(inspectors=[FakeMail("other")], selected=[FakeMail("another")])
        self.assertIsNone(server._open_in_outlook(outlook, FakeNamespace(), "mine"))

    def test_open_window_is_found_even_when_also_selected(self):
        outlook = FakeOutlook(inspectors=[FakeMail("x"), FakeMail("mine")], selected=[FakeMail("mine")])
        self.assertEqual(server._open_in_outlook(outlook, FakeNamespace(), "mine"), "window")

    def test_selected_in_the_main_window(self):
        outlook = FakeOutlook(inspectors=[], selected=[FakeMail("a"), FakeMail("mine")])
        self.assertEqual(server._open_in_outlook(outlook, FakeNamespace(), "mine"), "reading_pane")

    def test_unsaved_new_item_and_empty_inspector_are_skipped(self):
        outlook = FakeOutlook(inspectors=[FakeMail(""), None], selected=[])
        self.assertIsNone(server._open_in_outlook(outlook, FakeNamespace(), "mine"))

    def test_no_main_window(self):
        outlook = FakeOutlook(inspectors=[], selected=None)
        self.assertIsNone(server._open_in_outlook(outlook, FakeNamespace(), "mine"))


@unittest.skipUnless(pywintypes, "pywin32 is not installed")
class ComErrorWordingTests(unittest.TestCase):
    DISP_E_EXCEPTION = -2147352567

    def test_outlooks_own_sentence_leads(self):
        e = pywintypes.com_error(
            self.DISP_E_EXCEPTION, "Exception occurred.",
            (4096, "Microsoft Outlook", "This method can't be used with an inline response mail item.", None, 0,
             -2147467259),
            None,
        )
        self.assertEqual(format_com_error(e),
                         "This method can't be used with an inline response mail item. (Outlook error 0x80004005)")

    def test_no_detail_says_so(self):
        e = pywintypes.com_error(self.DISP_E_EXCEPTION, "Exception occurred.", None, None)
        text = format_com_error(e)
        self.assertTrue(text.startswith("COM Error (0x80020009): Exception occurred."), text)
        self.assertIn(NO_DETAIL_HINT, text)

    def test_other_com_errors_keep_the_old_form(self):
        e = pywintypes.com_error(-2147221005, "Invalid class string", None, None)
        self.assertEqual(format_com_error(e), "COM Error (0x800401F3): Invalid class string")


class NonComErrorTests(unittest.TestCase):
    def test_plain_exceptions_are_named(self):
        self.assertEqual(format_com_error(ValueError("Attachment not found: x")), "ValueError: Attachment not found: x")


if __name__ == "__main__":
    unittest.main()
