"""Timing check for the searches that were slow in real use. Read-only: it changes nothing.

Runs the server over stdio exactly as Claude would and times each call. Needs classic
Outlook running. The queries mirror three real sessions (a person search across accounts,
a company search in one account, and a company name given to contact_history across all
accounts, which took two minutes in Claude Desktop) so before/after numbers are comparable.

Run:  .venv\\Scripts\\python tests\\perf_check.py [--out results.json] [--account hbs.edu] [--person dan]
                                               [--company "Earl Gray"] [--domain earlgrayandsons.com]

Each call's wall time includes any wait behind a previous slow call on the COM thread,
which is what the user experiences.
"""
import argparse
import asyncio
import json
import os
import sys
import time
from datetime import datetime, timedelta

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
try:
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def log(msg):
    print(msg, file=sys.stderr, flush=True)


def summarize(result):
    """Short description of a tool result: hit count, error, partial flag."""
    if isinstance(result, str):
        text = result.strip()
        if text.lower().startswith("error"):
            return "ERROR: " + text[:110]
        return "text: " + text[:80]
    if isinstance(result, list):
        return f"{len(result)} hits"
    if isinstance(result, dict):
        if "error" in result:
            return "ERROR: " + str(result["error"])[:110]
        bits = []
        if "results" in result and isinstance(result["results"], list):
            bits.append(f"{len(result['results'])} hits")
        if "message_count" in result:
            bits.append(f"{result['message_count']} messages")
        if "received_count" in result:
            bits.append(f"received={result['received_count']} sent={result['sent_count']}")
            m = result.get("meetings") or {}
            if m:
                bits.append(f"meetings past={m.get('past_count')} next={'yes' if m.get('next') else 'no'}")
        if result.get("partial"):
            bits.append("PARTIAL")
        return ", ".join(bits) or "ok"
    return str(type(result).__name__)


async def run(account: str, person: str, out_path: str | None, company: str = "Earl Gray",
              domain: str = "earlgrayandsons.com"):
    from mcp.client.stdio import stdio_client, StdioServerParameters
    from mcp.client.session import ClientSession

    params = StdioServerParameters(command=sys.executable, args=["-m", "outlook_desktop_mcp"], cwd=ROOT)
    rows = []
    async with stdio_client(params) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()

            async def call(label, name, **args):
                t0 = time.perf_counter()
                try:
                    result = await session.call_tool(name, args)
                    text = result.content[0].text if result.content else ""
                    try:
                        parsed = json.loads(text)
                    except ValueError:
                        parsed = text
                except Exception as e:  # transport-level failure
                    parsed = f"Error: {type(e).__name__}: {e}"
                dt = time.perf_counter() - t0
                desc = summarize(parsed)
                rows.append({"label": label, "tool": name, "args": args, "seconds": round(dt, 2), "result": desc})
                log(f"{dt:7.2f}s  {label:<48} {desc}")
                return parsed

            health = await call("health_check", "health_check")
            if isinstance(health, dict):
                log(f"          server {health.get('server_version')}; accounts: "
                    f"{[a['display_name'] for a in health.get('accounts', [])]}")

            # --- the PennSpring session: one account, company name ---
            await call("search: sender=pennspring, inbox+subfolders", "search_emails",
                       query="", sender="pennspring", folder="inbox", include_subfolders=True, count=30, account=account)
            await call("search: 'PennSpring' in Sent Items", "search_emails",
                       query="PennSpring", folder="sent", count=20, account=account)

            # --- the Dan session: subject word in one account, then a person across all accounts ---
            hits = await call("search: 'Sanders' inbox+subfolders", "search_emails",
                              query="Sanders", folder="inbox", include_subfolders=True, count=10, account=account)
            await call(f"contact_history: '{person}' all accounts", "contact_history", person=person, account="all")
            await call("contact_history: 'sanjohninc' all accounts", "contact_history", person="sanjohninc", account="all")
            await call("contact_history: 'pennspring.com' all, 400 days", "contact_history",
                       person="pennspring.com", days=400, account="all")

            first = None
            if isinstance(hits, list) and hits:
                first = hits[0]
            elif isinstance(hits, dict) and hits.get("results"):
                first = hits["results"][0]
            if first:
                await call("get_conversation on first 'Sanders' hit", "get_conversation",
                           entry_id=first["entry_id"], count=12, max_body_chars=1500, account=account)

            await call("search_events: 'Sanders' all calendars May-Oct", "search_events",
                       query="Sanders", start_date="2026-05-01", end_date="2026-10-31", calendar="all", account="heartland")

            # --- the Earl Gray session: a company name, which lives in subjects, not in people fields.
            # Claude Desktop opened with contact_history (four sweeps, 53 s of server time, nothing
            # found) before a subject search hit. The same three shapes, so the fix is measurable.
            hist = await call(f"contact_history: '{company}' all accounts (company name)", "contact_history",
                              person=company, account="all")
            if isinstance(hist, dict) and hist.get("subject_matches"):
                log(f"          subject_matches: {len(hist['subject_matches'])} (fallback found the thread)")
            await call(f"contact_history: '{domain}' all accounts, 730 days", "contact_history",
                       person=domain, days=730, account="all")
            await call(f"search: '{company}' folder=all, all accounts (one call)", "search_emails",
                       query=company, folder="all", account="all", count=10)

    total = sum(r["seconds"] for r in rows)
    log(f"\nTOTAL {total:.1f}s over {len(rows)} calls")
    if out_path:
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump({"when": datetime.now().isoformat(timespec="seconds"), "account": account,
                       "person": person, "total_seconds": round(total, 1), "calls": rows}, f, indent=2)
        log(f"saved {out_path}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--account", default="hbs.edu", help="account substring for the single-account searches")
    ap.add_argument("--person", default="dan", help="needle for the all-account contact_history")
    ap.add_argument("--company", default="Earl Gray", help="company name for the contact_history-by-mistake case")
    ap.add_argument("--domain", default="earlgrayandsons.com", help="that company's email domain")
    ap.add_argument("--out", default=None, help="write a JSON summary here")
    a = ap.parse_args()
    asyncio.run(run(a.account, a.person, a.out, a.company, a.domain))
