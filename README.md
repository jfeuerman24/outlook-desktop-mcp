# outlook-desktop-mcp

[![PyPI](https://img.shields.io/pypi/v/outlook-desktop-mcp)](https://pypi.org/project/outlook-desktop-mcp/)
[![Python](https://img.shields.io/pypi/pyversions/outlook-desktop-mcp)](https://pypi.org/project/outlook-desktop-mcp/)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS-blue)]()

**Turn your running Outlook Desktop into an MCP server.** No Microsoft Graph API, no Entra app registration, no OAuth tokens — just your local Outlook and the authentication you already have.

Any MCP client (Claude Code, Claude Desktop, etc.) can then send emails, manage your calendar, create tasks, handle attachments, and more — all through your existing Outlook session.

> **New here? Follow [SETUP.md](SETUP.md)** — step-by-step instructions for the Claude desktop app
> and Claude Code, including the Outlook security prompt.
>
> **This is a fork** of [Aanerud/outlook-desktop-mcp](https://github.com/Aanerud/outlook-desktop-mcp)
> (AGPL-3.0, see [LICENSE](LICENSE)). It adds the changes below; everything else is the upstream project.

## What this fork adds

| Change | Where |
|---|---|
| **Read-only mode.** `outlook-desktop-mcp --read-only` (or `OUTLOOK_MCP_READ_ONLY=1`) registers only tools that read. Send, reply, move, mark, create, update, delete, respond, complete and toggle tools are not exposed at all, so a prompt (or a malicious email) cannot make the client act on your mailbox. Works on Windows and macOS. | `readonly.py`, `entrypoint.py` |
| **Every calendar, not just the default.** New `list_calendars` tool; `list_events` and `search_events` take `calendar` (`""` = main Calendar, `"all"` = every calendar in the account, or a name/path such as `Calendar/Family`). Each event says which calendar it came from. | `tools/calendars.py` |
| **Thread reconstruction.** New `get_conversation` tool walks Outlook's conversation across folders and dates, deduped by Message-ID, oldest first, with a ConversationTopic fallback. | `tools/conversation.py` |
| **Quoted history trimmed before the length cap.** `read_email` returns the message's own text (`body_trimmed` says if anything was cut); `include_history=true` gives the raw body. Outlook's "You don't often get email from…" banner is removed. | `utils/text.py` |
| **Sender filter and subtree search.** `list_emails` and `search_emails` take `sender` (substring of address or name) and `include_subfolders`. `search_emails` may be called with just a sender. | `server.py` |
| **Exchange addresses resolved to SMTP.** `sender` is a real address instead of `/o=ExchangeLabs/…`; full reads include `recipients` with resolved addresses. | `utils/addresses.py` |
| **ISO 8601 dates with the local UTC offset** everywhere (`2026-10-26T17:40:00-04:00`); Outlook's "no date" becomes `null`. Events that started before the range but overlap it are included. | `utils/dates.py` |
| **Outgoing mail in a fixed font.** `send_email` and `reply_email` write HTML in Aptos 12 pt black (overridable with `OUTLOOK_MCP_FONT_*`, or `OUTLOOK_MCP_PLAIN_TEXT=1`); replies are injected above Outlook's quoted history instead of replacing it with plain text. | `utils/html.py` |
| **Offline unit tests** (`python -m unittest discover -s tests`) and a GitHub Actions workflow that runs them. PyPI publishing is manual in this fork. | `tests/test_*.py` |

## Quick Start

**1. Install** (requires Python 3.12+). From this fork:

```bash
git clone https://github.com/jfeuerman24/outlook-desktop-mcp.git
cd outlook-desktop-mcp
python -m venv .venv
.venv\Scripts\activate
pip install -e .
```

(The upstream release is `pip install outlook-desktop-mcp`; it does not include this fork's changes.)

**2. Register with Claude Code:**

```bash
claude mcp add outlook-desktop -- C:\path\to\outlook-desktop-mcp\.venv\Scripts\outlook-desktop-mcp.exe --read-only
```

Drop `--read-only` if you want Claude to be able to send, move and create things.

**Or register with Claude Desktop:** Settings → Developer → Edit Config, then add to `claude_desktop_config.json` (use the full path; Claude Desktop starts servers with a minimal PATH):

```json
{
  "mcpServers": {
    "outlook-desktop": {
      "command": "C:\\path\\to\\outlook-desktop-mcp\\.venv\\Scripts\\outlook-desktop-mcp.exe",
      "args": ["--read-only"]
    }
  }
}
```

**3. Open Outlook and start a Claude session.** That's it — tools are available immediately.

## How It Works — Platform Routing

When the server starts, it checks which operating system it is running on and takes one of two paths:

```
                        outlook-desktop-mcp starts
                                  |
                          sys.platform check
                         /                  \
                   "win32"                "darwin"
                      |                      |
              ┌───────┴────────┐    ┌────────┴────────┐
              │  server.py     │    │  server_mac.py   │
              │  COM Bridge    │    │  AppleScript     │
              │  (29 tools)    │    │  Bridge          │
              │                │    │  (22 tools)      │
              └───────┬────────┘    └────────┬─────────┘
                      |                      |
              OUTLOOK.EXE via         Microsoft Outlook
              COM / STA thread        via osascript
                      |                      |
              Exchange / M365         Exchange / M365
```

**Both paths use your locally running Outlook app and its existing authenticated session.** No cloud credentials, no Graph API tokens — the server inherits whatever account Outlook is signed into.

### Why two paths?

Windows Outlook (Classic) exposes a rich COM automation interface — the Outlook Object Model (`MSOUTL.OLB`). This has been the standard way to programmatically control Outlook on Windows for over 20 years. It provides deep access to mail rules, categories, MAPI properties, and the full folder hierarchy.

Mac Outlook does not support COM. Instead, it exposes an AppleScript dictionary that can be driven via the `osascript` command. The AppleScript interface covers the core operations — email, calendar, tasks — but does not expose rules, categories, or certain advanced MAPI features. This is a limitation of what Microsoft chose to include in Outlook for Mac's scripting dictionary, not a limitation of this project.

The server is structured as two parallel implementations with identical tool names and signatures, so MCP clients see the same interface regardless of platform. Tools that are not available on a given platform are simply not registered.

## Requirements

### Windows

- **Outlook Desktop (Classic)** — the `OUTLOOK.EXE` that comes with Microsoft 365 / Office. The new "modern" Outlook (`olk.exe`) does **not** support COM
- **Python 3.12+** (x64 or ARM64)
- **Outlook must be running** when the MCP server starts

Both x64 and ARM64 Windows are supported. On ARM64, all dependencies (`pywin32`, `mcp`, `pydantic-core`, `cryptography`, `cffi`, `rpds-py`) have prebuilt `win_arm64` wheels — see the [ARM64 install notes](#arm64-windows) below for the one extra `pip` flag you need.

#### Outlook "Programmatic Access" security prompts

When the MCP server first touches Outlook's COM API, Outlook may show a dialog: *"A program is trying to access email address information stored in Outlook"* (the Object Model Guard / Programmatic Access prompt). This is Outlook's protection against malicious automation.

You have three options:

1. **Click "Allow access for 10 minutes"** every time you start a session. Fine for casual use.
2. **Get the antivirus status to "Valid"** in *File > Options > Trust Center > Trust Center Settings > Programmatic Access*. When that line reads `Valid`, the "Never warn me about suspicious activity" radio becomes selectable and the prompts go away. On most personal machines with current Defender this works out of the box.
3. **Apply the registry policy** in [`docs/suppress-outlook-oom-prompts.reg`](docs/suppress-outlook-oom-prompts.reg). From an **elevated** PowerShell or Command Prompt (Win+X → *Terminal (Admin)*), run:

   ```powershell
   reg import "C:\path\to\outlook-desktop-mcp\docs\suppress-outlook-oom-prompts.reg"
   ```

   Then fully quit Outlook (check Task Manager for stray `OUTLOOK.EXE` processes) and reopen it. This writes `AdminSecurityMode=3` and approves all `PromptOOM*` categories under `HKLM\Software\Policies\Microsoft\Office\16.0\Outlook\Security`, which Outlook honors regardless of AV status. On Intune/MDM-managed corporate devices, `HKCU\Software\Policies\...\Outlook` is locked and the HKLM keys may be overwritten on next policy sync — if `reg import` fails or the prompts come back, ask IT to push the equivalent settings via Group Policy.

> **Windows 11 ARM64 note:** Defender does not register with Outlook's `IOfficeAntiVirus` interface on ARM64, so Trust Center shows *"Antivirus status: Invalid"* and the *"Never warn me about suspicious activity"* radio stays greyed out **even when Outlook is launched as Administrator**. Option 2 is unavailable on ARM64; the `reg import` from option 3 is the only durable suppression path.

### macOS

- **Microsoft Outlook for Mac** — version 16.x or later
- **Python 3.12+**
- **Outlook must be running** when the MCP server starts

#### Required macOS permissions

The first time a tool runs, macOS will show **two permission prompts** that you must approve:

1. **Privacy & Automation** — a system dialog asks: *"python3.12 wants to control Microsoft Outlook"*. Click **Allow** to let the server send AppleScript commands to Outlook.

2. **Accessibility** — to read your Exchange/M365 inbox, the server uses macOS UI scripting (System Events). This requires Accessibility access for `python3.12`:
   - Open **System Settings > Privacy & Security > Accessibility**
   - Find **python3.12** in the list (it appears after the first prompt)
   - Toggle it **on**

   Without Accessibility enabled, calendar, tasks, and local folder tools will work, but listing Exchange inbox messages will return empty results.

Both permissions are one-time setup — macOS remembers them for future sessions.

## Available Tools by Platform

### Email

| Tool | Windows | macOS | Description |
|------|:-------:|:-----:|-------------|
| `send_email` | yes | yes | Send an email with To/CC/BCC, plain text or HTML body |
| `list_emails` | yes | yes | List recent emails from any folder, with optional unread, date, sender and subfolder filters |
| `read_email` | yes | yes | Read full email content by entry ID or subject search (quoted history trimmed unless `include_history`) |
| `search_emails` | yes | yes | Full-text search across email subjects and bodies, optionally by sender and across subfolders |
| `get_conversation` | yes | — | Every message of a thread, oldest first, each once, across folders and dates |
| `reply_email` | yes | yes | Reply or reply-all, preserving the conversation thread |
| `mark_as_read` | yes | yes | Mark a specific email as read |
| `mark_as_unread` | yes | yes | Mark a specific email as unread |
| `move_email` | yes | yes | Move an email to Archive, Trash, or any folder |
| `list_folders` | yes | yes | Browse the folder hierarchy with item counts |

### Calendar

| Tool | Windows | macOS | Description |
|------|:-------:|:-----:|-------------|
| `list_calendars` | yes | — | List every calendar folder in an account (Birthdays, holidays, shared, subscribed) |
| `list_events` | yes | yes | List events within a date range, from one calendar or `calendar="all"` |
| `get_event` | yes | yes | Read full event details by entry ID |
| `create_event` | yes | yes | Create a personal calendar appointment |
| `create_meeting` | yes | yes | Create a meeting and send invitations to attendees |
| `update_event` | yes | yes | Modify an existing event's subject, time, location, etc. |
| `delete_event` | yes | yes | Delete an appointment or cancel a meeting |
| `respond_to_meeting` | yes | — | Accept, decline, or tentatively accept a meeting invite |
| `search_events` | yes | yes | Search calendar events by keyword within a date range |

### Tasks

| Tool | Windows | macOS | Description |
|------|:-------:|:-----:|-------------|
| `list_tasks` | yes | yes | List pending or completed tasks, sorted by due date |
| `get_task` | yes | yes | Read full task details including body and completion status |
| `create_task` | yes | yes | Create a new task with subject, due date, importance |
| `complete_task` | yes | yes | Mark a task as complete |
| `delete_task` | yes | yes | Remove a task |

### Attachments

| Tool | Windows | macOS | Description |
|------|:-------:|:-----:|-------------|
| `list_attachments` | yes | yes | List all attachments on an email or calendar event |
| `save_attachment` | yes | yes | Download an attachment to a local directory |

### Categories, Rules, Out of Office (Windows only)

These tools rely on COM-specific APIs (MAPI property accessors, the Rules object model, and the Categories collection) that Outlook for Mac does not expose through AppleScript.

| Tool | Windows | macOS | Description |
|------|:-------:|:-----:|-------------|
| `list_categories` | yes | — | List all available color categories in Outlook |
| `set_category` | yes | — | Set or clear categories on any email, event, or task |
| `list_rules` | yes | — | List all mail rules with enabled/disabled status |
| `toggle_rule` | yes | — | Enable or disable a mail rule by name |
| `get_out_of_office` | yes | — | Check whether Out of Office auto-reply is on or off |

**Total: 31 tools on Windows, 22 tools on macOS.** In `--read-only` mode: 17 on Windows, 10 on macOS.

### Read-only mode

```bash
outlook-desktop-mcp --read-only
```

or set `OUTLOOK_MCP_READ_ONLY=1`. The server then registers only: `list_accounts`, `list_emails`,
`read_email`, `search_emails`, `get_conversation`, `list_folders`, `list_calendars`, `list_events`,
`get_event`, `search_events`, `list_tasks`, `get_task`, `list_attachments`, `save_attachment`
(writes to disk, never to the mailbox), `list_categories`, `list_rules`, `get_out_of_office`.
The write tools are not merely refused — they are never registered, so the client does not see them.

### Working with several calendars

```
list_calendars account="personal"
→ Calendar (default), Calendar/Birthdays, Calendar/United States holidays, Calendar/Kate Work

list_events start_date="2026-10-27" end_date="2026-10-29" calendar="all" account="personal"
→ events from all four, each tagged with its calendar, sorted by start
```

To check availability across accounts, call `list_events` with `calendar="all"` once per account and
compare the results.

## Architecture Details

### Windows: COM Bridge (`com_bridge.py`)

All Outlook COM operations run on a dedicated thread using the Single-Threaded Apartment (STA) model, as required by COM. The async MCP event loop dispatches tool calls to this thread via a queue and awaits results, keeping COM threading rules respected and the MCP protocol non-blocking.

```
MCP tool call (async)
  → bridge.call(func, args)
    → queued to STA thread
      → func(outlook, namespace, args) executes on COM thread
    → result returned via threading.Event
  → JSON response back to MCP client
```

Each tool's inner function receives the live `Outlook.Application` and `MAPI.Namespace` COM objects and works directly with the Outlook Object Model — `GetItemFromID`, `CreateItem`, `Items.Restrict` with DASL filters, and so on.

### macOS: AppleScript Bridge (`applescript_bridge.py`)

Each tool call builds an AppleScript string and executes it as a subprocess via `osascript`. There is no persistent connection — every call is stateless.

```
MCP tool call (async)
  → build AppleScript string
    → asyncio.create_subprocess_exec("osascript", "-e", script)
    → parse stdout text into structured data
  → JSON response back to MCP client
```

Each tool constructs a single AppleScript that fetches all needed data in one `osascript` call (no per-message subprocess loops). Results come back as delimited text, which the server parses into the same JSON structure the Windows server produces.

**Key differences from Windows:**

- Entry IDs on macOS are **numeric** (e.g. `42`), not hex strings. They identify items within their folder context.
- Folder references use AppleScript's **locale-independent keywords** (`inbox`, `sent items`, `drafts`, `deleted items`) rather than localized folder names.
- Search uses AppleScript's `whose` clause (e.g. `messages whose subject contains "query"`) instead of DASL filters.
- User input is escaped for safe embedding in AppleScript strings to prevent script injection.

## Install from Source

### Windows (x64)

```bash
git clone https://github.com/Aanerud/outlook-desktop-mcp.git
cd outlook-desktop-mcp
python -m venv .venv
.venv\Scripts\activate
pip install pywin32 "mcp[cli]" -e .
python .venv\Scripts\pywin32_postinstall.py -install
```

Register from source using the launcher script:

```bash
claude mcp add outlook-desktop -- powershell.exe -Command "& 'C:\path\to\outlook-desktop-mcp\outlook-desktop-mcp.cmd' mcp"
```

### Windows (ARM64)

The `[cli]` extra of `mcp` transitively pulls in `cryptography`, and pip's default resolver may pick a version that lacks a `win_arm64` wheel — which then fails to build because it requires a Rust toolchain plus OpenSSL. Install without the `cli` extra and force wheels-only resolution:

```powershell
git clone https://github.com/Aanerud/outlook-desktop-mcp.git
cd outlook-desktop-mcp
& "C:\Program Files\Python312-arm64\python.exe" -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install --only-binary=:all: pywin32 mcp
python -m pip install --no-deps -e .
python .venv\Scripts\pywin32_postinstall.py -install
```

The base `mcp` package is sufficient for running the stdio server — the `[cli]` extra is only needed for the `mcp` developer CLI tools (`mcp dev`, `mcp inspector`), which aren't used at runtime.

Register from source the same way as x64:

```bash
claude mcp add outlook-desktop -- powershell.exe -Command "& 'C:\path\to\outlook-desktop-mcp\outlook-desktop-mcp.cmd' mcp"
```

### macOS

```bash
git clone https://github.com/Aanerud/outlook-desktop-mcp.git
cd outlook-desktop-mcp
python3 -m venv .venv
source .venv/bin/activate
pip install "mcp[cli]" -e .
```

Register from source:

```bash
claude mcp add outlook-desktop -- /path/to/outlook-desktop-mcp/.venv/bin/python -m outlook_desktop_mcp
```

## Usage Examples

Once registered, just talk to Claude naturally:

- *"Show me my 10 most recent inbox emails"*
- *"Read the email from Taylor about MLADS"*
- *"Send an email to alice@example.com about the project update"*
- *"What's on my calendar this week?"*
- *"Create a meeting with bob@example.com tomorrow at 2pm for 30 minutes"*
- *"Save the attachment from that email to my Downloads folder"*
- *"Create a task to review the quarterly report, due Friday, high importance"*
- *"Mark that email as read and move it to archive"*

Windows-only examples:

- *"What categories do I have? Set this email to 'Follow-up'"*
- *"List my mail rules"*
- *"Am I set as Out of Office?"*

## Why Not Microsoft Graph?

| | Microsoft Graph | outlook-desktop-mcp |
|---|---|---|
| Entra app registration | Required | Not needed |
| Admin consent | Required for mail permissions | Not needed |
| OAuth token management | You handle refresh tokens | Not needed |
| Tenant configuration | Required | Not needed |
| Works offline / cached | No | Yes (reads from local cache) |
| Setup time | 30-60 minutes | 2 minutes |
| Auth requirement | **Your own OAuth flow** | **Outlook is open** |

## Project Structure

```
outlook-desktop-mcp/
  src/outlook_desktop_mcp/
    entrypoint.py            # Platform detection + --read-only flag → routes to correct server
    readonly.py              # Read-only switch and the list of write tools
    server.py                # Windows MCP server (31 tools, COM automation)
    server_mac.py            # macOS MCP server (22 tools, AppleScript)
    com_bridge.py            # Async-to-COM threading bridge (Windows)
    applescript_bridge.py    # Async osascript execution (macOS)
    tools/
      _folder_constants.py   # Outlook enums and constants (Windows)
      calendars.py           # Every calendar folder in a store; calendar argument resolution
      conversation.py        # Conversation tree walk + ConversationTopic fallback
    utils/
      formatting.py          # Email/event/task data extraction (Windows)
      text.py                # Quoted-history trimming, external-sender banner removal (pure)
      dates.py               # pywintypes → ISO 8601 with offset (pure)
      addresses.py           # Exchange DN → SMTP resolution, recipients
      errors.py              # COM error formatting (Windows)
      applescript_helpers.py # AppleScript escaping, date formatting (macOS)
  tests/
    test_text.py             # unit: trimming (no Outlook)
    test_dates.py            # unit: ISO dates (no Outlook)
    test_readonly.py         # unit: read-only mode hides write tools (no Outlook)
    live_readonly_check.py   # live: the fork's additions over stdio, read-only, sends nothing
    phase1_com_test.py       # Email COM validation
    phase3_mcp_test.py       # Email MCP test (sends a test email)
    calendar_com_test.py     # Calendar COM validation
    calendar_mcp_test.py     # Calendar MCP test
    extras_com_test.py       # Tasks/attachments/categories/rules/OOF COM test
    extras_mcp_test.py       # Tasks/attachments/categories/rules/OOF MCP test
  outlook-desktop-mcp.cmd   # Windows launcher script
  pyproject.toml
```

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for the branching strategy and development setup.

## License

See [LICENSE](LICENSE) file.
