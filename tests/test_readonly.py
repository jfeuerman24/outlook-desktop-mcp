"""Read-only mode hides every write tool. Imports the server in a subprocess so the
environment variable is read fresh each time; no Outlook is contacted (the COM bridge
only connects when the server starts)."""
import json
import os
import subprocess
import sys
import unittest

from outlook_desktop_mcp.readonly import ENV_VAR, WRITE_TOOLS

_LIST_TOOLS = (
    "import asyncio, json; from outlook_desktop_mcp import server; "
    "print(json.dumps(sorted(t.name for t in asyncio.run(server.mcp.list_tools()))))"
)


def _tool_names(read_only: bool) -> list[str]:
    env = {k: v for k, v in os.environ.items() if k != ENV_VAR}
    if read_only:
        env[ENV_VAR] = "1"
    out = subprocess.run([sys.executable, "-c", _LIST_TOOLS], capture_output=True, text=True, env=env, timeout=120)
    assert out.returncode == 0, out.stderr[-2000:]
    return json.loads(out.stdout.strip().splitlines()[-1])


class ReadOnlyTests(unittest.TestCase):
    def test_full_mode_registers_everything(self):
        names = set(_tool_names(read_only=False))
        self.assertTrue(WRITE_TOOLS <= names, WRITE_TOOLS - names)
        for tool in ("list_calendars", "get_conversation", "list_emails", "search_emails", "read_email"):
            self.assertIn(tool, names)

    def test_read_only_mode_hides_write_tools(self):
        names = set(_tool_names(read_only=True))
        self.assertFalse(WRITE_TOOLS & names, WRITE_TOOLS & names)
        for tool in ("list_accounts", "list_calendars", "list_events", "get_conversation", "save_attachment"):
            self.assertIn(tool, names)


if __name__ == "__main__":
    unittest.main()
