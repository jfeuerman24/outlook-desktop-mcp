"""
Live check of how the server writes draft text, read back through Word (Outlook's editor):
blank lines are truly empty paragraphs, numbered items keep their numbers across blank lines,
and inline answers line up with the text of the numbered item they answer.

It WRITES, but only to two drafts of its own: a plain one, and one holding a Word-style numbered
list that the inline answers are written into. Each is checked by EntryID and subject before
anything is read or deleted, both are deleted at the end, and the main window's folder and
selection are never touched.

Run: .venv\\Scripts\\python tests\\inline_live_check.py
"""
import asyncio
import json
import os
import re
import sys
import time

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SUBJECT = "[outlook-mcp inline check] delete me"
try:
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

LIST = "margin-left:.5in;text-indent:-.25in"
SUB = "margin-left:1.0in;text-indent:-.25in"
SOURCE = (
    '<p class=MsoNormal>Jake, open points below.</p>'
    f'<p class=MsoNormal style="{LIST}">19.&nbsp;&nbsp;Restricted Business is redefined. Confirm reject.</p>'
    f'<p class=MsoNormal style="{SUB}">(a)&nbsp;&nbsp;Sub point about customers.</p>'
    f'<p class=MsoNormal style="{LIST}">20.&nbsp;&nbsp;Term is two years. Confirm ok.</p>'
    '<p class=MsoNormal>Thanks, Mitch</p>'
)


def log(msg):
    print(msg, file=sys.stderr, flush=True)


def check(cond, what):
    log(("  ok   " if cond else "  FAIL ") + what)
    if not cond:
        raise SystemExit(f"live check failed: {what}")


class Own:
    """COM access limited to items this script created."""

    def __init__(self):
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.ns = self.outlook.GetNamespace("MAPI")
        self.ids: list[str] = []

    def item(self, entry_id):
        it = self.ns.GetItemFromID(entry_id)
        if (entry_id not in self.ids or not self.ns.CompareEntryIDs(str(it.EntryID), entry_id)
                or SUBJECT not in str(it.Subject)):
            raise SystemExit("refusing to touch an item that is not the check's own")
        return it

    def make_source(self) -> str:
        mail = self.ns.GetDefaultFolder(16).Items.Add(0)
        mail.Subject = SUBJECT
        mail.To = "live-check@example.com"
        mail.HTMLBody = "<html><body>" + SOURCE + "</body></html>"
        mail.Save()
        entry_id = str(mail.EntryID)
        self.ids.append(entry_id)
        return entry_id

    def paragraphs(self, entry_id) -> list[dict]:
        """The draft's paragraphs as Word holds them, down to the quoted header or signature."""
        it = self.item(entry_id)
        doc = it.GetInspector.WordEditor
        out = []
        for i in range(1, min(doc.Paragraphs.Count, 40) + 1):
            p = doc.Paragraphs(i)
            out.append({"text": str(p.Range.Text), "left": float(p.LeftIndent), "first": float(p.FirstLineIndent),
                        "before": float(p.SpaceBefore), "after": float(p.SpaceAfter),
                        "number": str(p.Range.ListFormat.ListString)})
        it.Close(1)  # olDiscard: nothing was changed
        return out

    def cleanup(self):
        for entry_id in self.ids:
            try:
                self.item(entry_id).Delete()
            except Exception as e:
                log(f"cleanup: could not delete ..{entry_id[-8:]}: {e}")
        log(f"\ncleanup: {len(self.ids)} draft(s) of the check's own moved to Deleted Items")


def find(paras, needle):
    for i, p in enumerate(paras):
        if needle in p["text"]:
            return i
    raise SystemExit(f"live check failed: no paragraph containing {needle!r}")


async def run():
    from mcp.client.stdio import stdio_client, StdioServerParameters
    from mcp.client.session import ClientSession

    own = Own()
    env = {**os.environ, "PYTHONPATH": os.path.join(ROOT, "src"), "OUTLOOK_MCP_SIGNOFF_BLANK_LINES": "2"}
    params = StdioServerParameters(command=sys.executable, args=["-m", "outlook_desktop_mcp"], cwd=ROOT, env=env)
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

                log("draft_email: three paragraphs and a numbered list with blank lines between items")
                r = await call("draft_email", to="live-check@example.com", subject=SUBJECT,
                               body="First paragraph.\n\nSecond paragraph.\n\n1. One\n\n2. Two\n\n3. Three\n\n"
                                    "Last paragraph.\n\nBest,\nJake")
                check(isinstance(r, dict) and r.get("status") == "draft_saved", "draft saved")
                plain = r["entry_id"]
                own.ids.append(plain)
                paras = own.paragraphs(plain)
                z = find(paras, "Last paragraph")
                check([p["text"] for p in paras[z + 1:z + 3]] == ["\r", "\r"] and paras[z + 3]["text"].startswith("Best,"),
                      "exactly two empty lines between the last paragraph and the sign-off")
                check("Jake" in paras[z + 3]["text"], f"the name is on the line under it: {paras[z + 3]['text']!r}")
                r = await call("update_draft", entry_id=plain, body="Changed text.\n\n\n\n\nBest regards,\nJake")
                check(isinstance(r, dict) and r.get("status") == "draft_saved", "update_draft saved")
                after = own.paragraphs(plain)
                c = find(after, "Changed text")
                check([p["text"] for p in after[c + 1:c + 3]] == ["\r", "\r"] and after[c + 3]["text"].startswith("Best regards,"),
                      "update_draft: four blank lines typed become exactly two")
                check(not any("Last paragraph" in p["text"] for p in after), "update_draft replaced only the note")
                a, z = find(paras, "First paragraph"), find(paras, "Last paragraph")
                note = paras[a:z + 1]
                for p in note:
                    log(f"       {p['number']:>3} {p['text']!r}")
                blanks = [p for p in note if not p["text"].strip("\r\n\x07 \xa0")]
                check(len(blanks) == 5, f"five blank lines between the six blocks (got {len(blanks)})")
                check(all(p["text"] == "\r" for p in blanks), "every blank line is an empty paragraph, no character on it")
                check([p["number"] for p in note if p["number"]] == ["1.", "2.", "3."], "items are numbered 1, 2, 3")
                check(all(p["before"] == 0 and p["after"] == 0 for p in note), "no paragraph spacing anywhere in the note")

                # Outlook will not Reply to an unsent draft, so the answers go into the check's own
                # draft through update_draft, which writes them with the same code as draft_reply.
                log("\ninline answers under Word-style numbered items")
                reply = own.make_source()
                answers = [{"after": "Confirm reject", "text": "JF: Reject."},
                           {"after": "Sub point about customers", "text": "JF: Fine.\n\nSecond paragraph of the answer."},
                           {"after": "Term is two years", "text": "JF: Ok."}]
                r = await call("update_draft", entry_id=reply, inline_replies=json.dumps(answers))
                check(isinstance(r, dict) and r.get("status") == "draft_saved", f"answers saved: {str(r)[:120]}")
                paras = own.paragraphs(reply)
                start = find(paras, "19.")
                for p in paras[start:find(paras, "Thanks, Mitch") + 1]:
                    log(f"       left={p['left']:>5} first={p['first']:>6} {p['text']!r}")

                def answer(needle, item_needle):
                    i, item = find(paras, needle), paras[find(paras, item_needle)]
                    p = paras[i]
                    check(p["left"] == item["left"] and p["first"] == 0,
                          f"{needle!r} starts where the item's text starts ({p['left']} pt), no hanging indent")
                    check(p["before"] == 0 and p["after"] == 0, f"{needle!r} has no paragraph spacing")
                    return i

                i = answer("JF: Reject.", "19.")
                check(paras[i + 1]["text"] == "\r", "one empty line follows the answer")
                check(paras[i + 2]["text"].startswith("(a)"), "then the next item, directly")
                i = answer("JF: Fine.", "(a)")
                check(paras[i + 1]["text"] == "\r" and "Second paragraph" in paras[i + 2]["text"],
                      "a blank line inside an answer is kept")
                check(paras[i + 2]["left"] == paras[find(paras, "(a)")]["left"], "the second paragraph is indented too")
                check(paras[i + 3]["text"] == "\r" and paras[i + 4]["text"].startswith("20."), "one empty line, then item 20")
                i = answer("JF: Ok.", "20.")
                check(paras[i + 1]["text"] == "\r", "one empty line follows the last answer")

                log("\nupdate_draft replaces the answers without leaving old ones or their blank lines behind")
                r = await call("update_draft", entry_id=reply, inline_replies=json.dumps(answers[:1]))
                check(isinstance(r, dict) and r.get("status") == "draft_saved", "updated")
                paras = own.paragraphs(reply)
                texts = [p["text"] for p in paras]
                check(not any("JF: Fine." in t or "JF: Ok." in t for t in texts), "the two dropped answers are gone")
                i = find(paras, "(a)")
                check(paras[i + 1]["text"].startswith("20."), "and so are their blank lines")
    finally:
        time.sleep(0.5)
        own.cleanup()
    log("\nOK")


if __name__ == "__main__":
    asyncio.run(run())
