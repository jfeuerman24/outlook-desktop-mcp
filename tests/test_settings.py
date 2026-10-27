"""Unit tests for utils.settings — env > file > default. No Outlook needed."""
import json
import os
import pathlib
import tempfile
import unittest
from unittest import mock

from outlook_desktop_mcp.utils import settings


class SettingsTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = pathlib.Path(self.tmp.name) / "settings.json"
        self.env = mock.patch.dict(os.environ, {"OUTLOOK_MCP_SETTINGS_FILE": str(self.path)}, clear=False)
        self.env.start()
        for spec in settings.SETTINGS.values():
            os.environ.pop(spec["env"], None)

    def tearDown(self):
        self.env.stop()
        self.tmp.cleanup()

    def test_defaults_when_nothing_set(self):
        self.assertEqual(settings.get_with_source("default_account"), ("", "default"))
        self.assertEqual(settings.get_with_source("open_drafts"), (False, "default"))
        self.assertEqual(settings.get_with_source("open_in_outlook"), ("meetings", "default"))
        self.assertEqual(settings.get("hide_accounts"), [])
        self.assertEqual(settings.get("font_family"), "Aptos")

    def test_open_in_outlook_default_opens_meetings_only(self):
        self.assertEqual(settings.open_in_outlook_mode(), "meetings")
        self.assertTrue(settings.should_open("meeting"))
        for kind in ("email", "reply", "reply_all", "forward", "update"):
            self.assertFalse(settings.should_open(kind), kind)

    def test_open_in_outlook_is_a_validated_choice(self):
        self.assertEqual(settings.set_value("open_in_outlook", " ALL "), "all")
        self.assertTrue(settings.should_open("email"))
        self.assertEqual(settings.set_value("open_in_outlook", "none"), "none")
        self.assertFalse(settings.should_open("meeting"))
        with self.assertRaises(ValueError):
            settings.set_value("open_in_outlook", "sometimes")
        # A bad value already in the env or file falls back to the default instead of breaking lookups.
        with mock.patch.dict(os.environ, {"OUTLOOK_MCP_OPEN_IN_OUTLOOK": "bogus"}):
            self.assertEqual(settings.get("open_in_outlook"), "meetings")

    def test_legacy_open_drafts_still_honoured_until_the_new_setting_is_set(self):
        settings.set_value("open_drafts", "true")
        self.assertEqual(settings.open_in_outlook_mode(), "all")
        settings.set_value("open_drafts", "false")
        self.assertEqual(settings.open_in_outlook_mode(), "none")
        settings.set_value("open_in_outlook", "meetings")
        self.assertEqual(settings.open_in_outlook_mode(), "meetings")
        with mock.patch.dict(os.environ, {"OUTLOOK_MCP_OPEN_DRAFTS": "1"}):
            self.assertEqual(settings.open_in_outlook_mode(), "meetings")

    def test_set_value_writes_file_and_applies_immediately(self):
        self.assertEqual(settings.set_value("default_account", "Jake@heartland-gp.com"), "Jake@heartland-gp.com")
        self.assertTrue(self.path.exists())
        self.assertEqual(json.loads(self.path.read_text())["default_account"], "Jake@heartland-gp.com")
        self.assertEqual(settings.get_with_source("default_account"), ("Jake@heartland-gp.com", "file"))

    def test_list_and_bool_coercion(self):
        self.assertEqual(settings.set_value("hide_accounts", "Pickleball; Jake Feuerman ;"), ["Pickleball", "Jake Feuerman"])
        self.assertEqual(settings.get("hide_accounts"), ["Pickleball", "Jake Feuerman"])
        self.assertIs(settings.set_value("open_drafts", "yes"), True)
        self.assertIs(settings.set_value("open_drafts", "0"), False)
        self.assertIs(settings.get("open_drafts"), False)

    def test_env_overrides_file(self):
        settings.set_value("font_family", "Calibri")
        with mock.patch.dict(os.environ, {"OUTLOOK_MCP_FONT_FAMILY": "Georgia"}):
            self.assertEqual(settings.get_with_source("font_family"), ("Georgia", "env"))
        self.assertEqual(settings.get_with_source("font_family"), ("Calibri", "file"))

    def test_clear_value_restores_default(self):
        settings.set_value("font_size_pt", "11")
        settings.clear_value("font_size_pt")
        self.assertEqual(settings.get_with_source("font_size_pt"), ("12", "default"))
        self.assertNotIn("font_size_pt", json.loads(self.path.read_text()))

    def test_unknown_key_rejected(self):
        with self.assertRaises(KeyError):
            settings.set_value("nope", "x")

    def test_external_edit_is_picked_up(self):
        settings.set_value("default_account", "A")
        data = json.loads(self.path.read_text())
        data["default_account"] = "B"
        self.path.write_text(json.dumps(data))
        os.utime(self.path, (self.path.stat().st_atime, self.path.stat().st_mtime + 5))
        self.assertEqual(settings.get("default_account"), "B")

    def test_all_settings_shape(self):
        info = settings.all_settings()
        self.assertEqual(set(info), set(settings.SETTINGS))
        self.assertEqual(info["open_drafts"]["env_var"], "OUTLOOK_MCP_OPEN_DRAFTS")
        self.assertEqual(info["open_in_outlook"]["env_var"], "OUTLOOK_MCP_OPEN_IN_OUTLOOK")
        self.assertIn("source", info["default_account"])


if __name__ == "__main__":
    unittest.main()
