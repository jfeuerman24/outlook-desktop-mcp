"""
Live check of the draft handling added in 0.11.2, over stdio against the running classic
Outlook. Needs Outlook open.

It WRITES, but only to one item of its own: a draft "[outlook-mcp live check] delete me" in the
default account's Drafts folder, which it opens in an Outlook window for a few seconds, updates,
closes, moves to Deleted Items and then deletes for good. Every write through COM is preceded by
an EntryID comparison against the id the script created, and it never touches the main window's
folder or selection.

Run: .venv\\Scripts\\python tests\\drafts_live_check.py
"""
import asyncio
import json
import os
import sys
import tempfile
import time

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SUBJECT = "[outlook-mcp live check] delete me"
try:
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def log(msg):
    print(msg, file=sys.stderr, flush=True)


def check(cond, what):
    log(("  ok   " if cond else "  FAIL ") + what)
    if not cond:
        raise SystemExit(f"live check failed: {what}")


class OwnItem:
    """COM access to the one item this script created; refuses to act on anything else."""

    def __init__(self, entry_id: str):
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.ns = self.outlook.GetNamespace("MAPI")
        self.entry_id = entry_id

    def _mine(self, other_id) -> bool:
        return bool(self.ns.CompareEntryIDs(str(other_id), self.entry_id))

    def item(self):
        it = self.ns.GetItemFromID(self.entry_id)
        if not self._mine(it.EntryID) or str(it.Subject) != SUBJECT:
            raise SystemExit("refusing to touch an item that is not the live check's own")
        return it

    def display(self):
        self.item().Display()
        time.sleep(1.5)

    def inspector(self):
        inspectors = self.outlook.Inspectors
        for i in range(1, inspectors.Count + 1):
            insp = inspectors.Item(i)
            try:
                cur = insp.CurrentItem
                if cur is not None and self._mine(cur.EntryID):
                    return insp
            except Exception:
                continue
        return None

    def close_window(self):
        insp = self.inspector()
        if insp is not None:
            insp.Close(0)  # olSave
            time.sleep(1)

    def delete_for_good(self):
        it = self.item()
        it.Delete()


async def run():
    from mcp.client.stdio import stdio_client, StdioServerParameters
    from mcp.client.session import ClientSession

    tmp = tempfile.mkdtemp(prefix="outlook-mcp-live-")
    a = os.path.join(tmp, "live check a.txt")
    b = os.path.join(tmp, "live check b.txt")
    for p, data in ((a, "a"), (b, "bb")):
        with open(p, "w") as f:
            f.write(data)
    missing = os.path.join(tmp, "not-here.pdf")

    env = {**os.environ, "PYTHONPATH": os.path.join(ROOT, "src")}
    params = StdioServerParameters(command=sys.executable, args=["-m", "outlook_desktop_mcp"], cwd=ROOT, env=env)
    own = None
    new_id = ""
    try:
        async with stdio_client(params) as (read, write):
            async with ClientSession(read, write) as session:
                await session.initialize()

                async def call(name, **args):
                    result = await session.call_tool(name, args)
                    text = result.content[0].text
                    try:
                        return json.loads(text)
                    except ValueError:
                        return text

                health = await call("health_check")
                log(f"server {health.get('server_version')} / Outlook {health.get('outlook_version')}")
                check(health.get("server_version") == "0.11.2", "health_check reports 0.11.2")

                log("\ndraft_email with a missing attachment path")
                r = await call("draft_email", to="live-check@example.com", subject=SUBJECT, body="never saved",
                               attachments=missing)
                check(isinstance(r, str) and "Attachment not found" in r, f"refused up front: {str(r)[:90]}")

                log("\ndraft_email with one attachment")
                r = await call("draft_email", to="live-check@example.com", subject=SUBJECT,
                               body="Live check of outlook-desktop-mcp 0.11.2. Delete me.", attachments=a)
                check(isinstance(r, dict) and r.get("status") == "draft_saved", "draft saved")
                check(len(r.get("files_attached", [])) == 1 and "files_failed" not in r, "one file attached, none failed")
                entry_id = r["entry_id"]
                own = OwnItem(entry_id)

                log("\nadd_attachment")
                r = await call("add_attachment", entry_id=entry_id, paths=b)
                check(r.get("status") == "attached" and r.get("attachment_count") == 2, "second file attached")
                r = await call("add_attachment", entry_id=entry_id, paths="")
                check("error" in r, "empty paths is a clean error")

                log("\nopen the draft in an Outlook window (the script's own draft)")
                own.display()
                check(own.inspector() is not None, "window is open")

                r = await call("move_email", entry_id=entry_id, target_folder="deleted")
                check(isinstance(r, str) and r.startswith("Error: this email is open in an Outlook window"),
                      f"move refused while open: {str(r)[:80]}")

                r = await call("update_draft", entry_id=entry_id, body="Changed by the live check.")
                check(r.get("status") == "draft_saved" and r.get("open_in_outlook") == "window",
                      "update_draft saved and reports open_in_outlook=window")
                insp = own.inspector()
                check(insp is not None, "the window is still open afterwards")
                shown = str(insp.CurrentItem.HTMLBody) if insp is not None else ""
                check("Changed by the live check" in shown, "the window shows the new text")

                own.close_window()
                check(own.inspector() is None, "window closed")

                log("\nmove_email after the window is closed")
                r = await call("move_email", entry_id=entry_id, target_folder="deleted")
                check(isinstance(r, dict) and r.get("status") == "moved" and r.get("entry_id"),
                      "moved; new entry_id returned")
                new_id = r["entry_id"]
                r = await call("read_email", entry_id=new_id)
                check(isinstance(r, dict) and "Deleted" in str(r.get("folder", "")), f"readable at the new id in {r.get('folder')}")
                check(r.get("attachment_count") == 2, "both attachments travelled")
    finally:
        if own is not None and new_id:
            own.entry_id = new_id
            own.delete_for_good()
            log("\ncleanup: the live check's item was deleted from Deleted Items")
        elif own is not None:
            try:
                own.close_window()
                own.item().Delete()
                log("\ncleanup: the live check's draft was deleted")
            except Exception as e:
                log(f"\ncleanup: could not delete the draft: {e}")
        for p in (a, b):
            try:
                os.remove(p)
            except OSError:
                pass
        try:
            os.rmdir(tmp)
        except OSError:
            pass
    log("\nOK")


if __name__ == "__main__":
    asyncio.run(run())
