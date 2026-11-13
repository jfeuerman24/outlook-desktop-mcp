"""
Live check of date filters. READ-ONLY: it counts and lists mail, and writes nothing.

1. For a run of start times (this morning's small hours, the last few hours, and yesterday
   evening, whose UTC equivalents are small hours too), Outlook must count the same Inbox mail
   under the server's Jet filter (local time) and its DASL filter (UTC).
2. Over stdio, search_emails with a start_date must return the newest Inbox mail, the mail a
   date filter used to hide for as many hours as the PC is from UTC.

Run: .venv\\Scripts\\python tests\\dates_live_check.py
"""
import asyncio
import json
import os
import sys
from datetime import datetime, timedelta

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def log(msg):
    print(msg, file=sys.stderr, flush=True)


def check(cond, what):
    log(("  ok   " if cond else "  FAIL ") + what)
    if not cond:
        raise SystemExit(f"live check failed: {what}")


def filters_agree():
    import pythoncom
    import win32com.client
    from outlook_desktop_mcp.utils import dasl

    pythoncom.CoInitialize()
    inbox = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI").GetDefaultFolder(6)
    now = datetime.now().replace(second=0, microsecond=0)
    midnight = now.replace(hour=0, minute=0)
    starts = [midnight + timedelta(hours=h) for h in (1, 4, 7, 9)]
    starts += [now - timedelta(hours=h) for h in (1, 3, 6)]
    starts += [midnight - timedelta(hours=h) for h in (1, 3, 5, 30)]
    log("Jet (local) and DASL (UTC) count the same mail from each start time")
    for start in sorted(starts):
        jet = inbox.Items.Restrict(f"[ReceivedTime] >= '{dasl.jet_date(start)}'").Count
        utc = inbox.Items.Restrict(dasl.sql(dasl.received_since(start))).Count
        check(jet == utc, f"since {start:%a %m/%d %H:%M}: jet={jet} dasl={utc}")
    newest = inbox.Items
    newest.Sort("[ReceivedTime]", True)
    return newest.GetFirst()


async def run():
    from mcp.client.stdio import stdio_client, StdioServerParameters
    from mcp.client.session import ClientSession

    first = filters_agree()
    newest_id = str(first.EntryID) if first is not None else ""

    env = {**os.environ, "PYTHONPATH": os.path.join(ROOT, "src")}
    params = StdioServerParameters(command=sys.executable, args=["-m", "outlook_desktop_mcp"], cwd=ROOT, env=env)
    async with stdio_client(params) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()

            async def call(name, **args):
                text = (await session.call_tool(name, args)).content[0].text
                try:
                    return json.loads(text)
                except ValueError:
                    return text

            log("\nsearch_emails and list_emails agree on the default inbox for each window")
            today = datetime.now().strftime("%Y-%m-%d")
            yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
            windows = [{"start_date": today}, {"start_date": today + " 07:00"},
                       {"start_date": yesterday, "end_date": today},
                       {"start_date": yesterday, "end_date": yesterday}]
            for window in windows:
                listed = await call("list_emails", folder="inbox", count=200, **window)
                check(isinstance(listed, list), f"{window}: list_emails returned {len(listed)} message(s)")
                # search_emails needs something to search for: ask for each sender in the window
                # and expect every message list_emails showed from that sender, and no others.
                found, senders = [], sorted({m["sender"] for m in listed if m.get("sender")})
                for sender in senders:
                    hits = await call("search_emails", query="", sender=sender, folder="inbox", count=200, **window)
                    check(isinstance(hits, list), f"search sender={sender[:30]!r}: {len(hits) if isinstance(hits, list) else hits}")
                    found += [m for m in hits if m["entry_id"] not in {f["entry_id"] for f in found}]
                a, b = {m["entry_id"] for m in listed}, {m["entry_id"] for m in found}
                check(a == b, f"{window}: list_emails {len(a)} = search_emails by sender {len(b)}")
                if window == {"start_date": today} and listed:
                    times = sorted(m["received_time"] for m in found)
                    log(f"       newest found: {times[-1]}")
                    if newest_id and any(m["received_time"][:10] == today for m in listed):
                        check(newest_id in b, "search_emails returns the newest mail in the Inbox")
                if window.get("end_date") == yesterday:
                    check(all(m["received_time"][:10] == yesterday for m in found) and (found or not listed),
                          "an end date with no time covers that whole day and nothing after")
    log("\nOK")


if __name__ == "__main__":
    asyncio.run(run())
