"""
Live check of the fork's additions over stdio, in READ-ONLY mode.
Needs Outlook Desktop (Classic) running. Sends nothing, changes nothing.

Run: .venv\\Scripts\\python tests\\live_readonly_check.py [sender-substring]
"""
import asyncio
import json
import os
import sys
from datetime import datetime, timedelta

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
try:
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def log(msg):
    print(msg, file=sys.stderr, flush=True)


async def run(sender_needle: str):
    from mcp.client.stdio import stdio_client, StdioServerParameters
    from mcp.client.session import ClientSession

    params = StdioServerParameters(
        command=sys.executable, args=["-m", "outlook_desktop_mcp", "--read-only"], cwd=ROOT,
    )
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

            tools = sorted(t.name for t in (await session.list_tools()).tools)
            log(f"tools ({len(tools)}): {', '.join(tools)}")
            writes = {"send_email", "reply_email", "move_email", "delete_event", "create_task"}
            assert not (writes & set(tools)), "write tools leaked into read-only mode"

            accounts = await call("list_accounts")
            log(f"\naccounts: {[a['display_name'] for a in accounts]}")

            for acct in accounts:
                name = acct["display_name"]
                cals = await call("list_calendars", account=name)
                if isinstance(cals, list):
                    log(f"  calendars in {name}: " + "; ".join(
                        f"{c['path']}{' (default)' if c['is_default'] else ''}" for c in cals))

            today = datetime.now().strftime("%Y-%m-%d")
            end = (datetime.now() + timedelta(days=3)).strftime("%Y-%m-%d")
            for acct in accounts:
                name = acct["display_name"]
                events = await call("list_events", start_date=today, end_date=end, count=50, calendar="all", account=name)
                if isinstance(events, list):
                    log(f"\nevents {today}..{end} in {name} (all calendars): {len(events)}")
                    for e in events[:6]:
                        if "error" in e:
                            log(f"   [{e.get('calendar')}] unreadable: {e['error']}")
                        else:
                            log(f"   {e['start']} -> {e['end']}  [{e.get('calendar')}]  {e['subject'][:50]}")
                else:
                    log(f"  list_events error for {name}: {events}")

            since = (datetime.now() - timedelta(days=2)).strftime("%Y-%m-%d")
            found = None
            for acct in accounts:
                name = acct["display_name"]
                hits = await call("search_emails", query="", sender=sender_needle, folder="inbox",
                                  include_subfolders=True, start_date=since, count=5, account=name)
                if isinstance(hits, list) and hits:
                    log(f"\nsearch sender='{sender_needle}' in {name}: {len(hits)} hit(s)")
                    for h in hits:
                        log(f"   {h['received_time']}  {h['sender']:<38} {h['folder']:<24} {h['subject'][:45]}")
                    found = (name, hits[0])
                    break
            if not found:
                log(f"\nno emails from '{sender_needle}' in the last 2 days; using newest inbox mail")
                name = accounts[0]["display_name"]
                hits = await call("list_emails", count=1, account=name)
                found = (name, hits[0])

            name, first = found
            full = await call("read_email", entry_id=first["entry_id"], account=name)
            log(f"\nread_email: trimmed={full['body_trimmed']} truncated={full['body_truncated']} "
                f"sent={full['sent_time']} recipients={len(full['recipients'])}")
            log("   body: " + full["body"][:160].replace("\n", " | "))

            conv = await call("get_conversation", entry_id=first["entry_id"], include_bodies=False, account=name)
            log(f"\nget_conversation: source={conv['source']} messages={conv['message_count']} topic={conv['topic'][:50]}")
            for m in conv["messages"]:
                flag = "*" if m["is_requested_message"] else " "
                log(f"  {flag} {m['received_time']}  {m['sender']:<34} {m['folder']:<22} {m['subject'][:40]}")
            log("\nOK")


if __name__ == "__main__":
    asyncio.run(run(sys.argv[1] if len(sys.argv) > 1 else "freestone"))
