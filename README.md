# outlook-desktop-mcp

[![PyPI](https://img.shields.io/pypi/v/outlook-desktop-mcp)](https://pypi.org/project/outlook-desktop-mcp/)
[![Python](https://img.shields.io/pypi/pyversions/outlook-desktop-mcp)](https://pypi.org/project/outlook-desktop-mcp/)
[![Platform](https://img.shields.io/badge/platform-Windows-blue)]()

**Lets Claude use the Outlook that is already open on your Windows PC.**

Claude can look through your mail, calendars and tasks, find things, pull a whole thread
together, and write emails, replies and meeting invitations for you. **It never sends
anything.** Whatever it writes goes to your Drafts folder (or onto your calendar with the
invitations unsent) and pops up in Outlook for you to read and send yourself.

There is no Microsoft account setup, nothing for an admin to approve, and no password to hand
over. The server talks to the Outlook program on your computer the way an add-in would, so it
has the access you already have, and only while Outlook is open.

| Claude can | Claude cannot |
|---|---|
| List, search and read mail, calendars and tasks. Save attachments to a folder. Write draft emails and replies. Prepare meeting invitations. Create and change appointments and tasks. Move and mark mail, set categories, switch rules on and off. | Send anything. Accept or decline invitations. Cancel meetings you organised. |

> **This is a fork** of [Aanerud/outlook-desktop-mcp](https://github.com/Aanerud/outlook-desktop-mcp)
> (AGPL-3.0, see [LICENSE](LICENSE)). It is Windows-only and never sends; the changes are listed
> under [What this fork changes](#what-this-fork-changes).

## Getting started (Claude does the setup)

You need a Windows PC with **classic Outlook** signed in, and the **Claude desktop app**. Allow
five minutes. You will not type any commands; Claude runs them and explains as it goes.

1. **Make an empty folder** for it, for example `C:\Users\<you>\outlook-desktop-mcp`
   (open your user folder in File Explorer, right-click → **New** → **Folder**). Avoid OneDrive folders.
2. **Open it in Claude.** In the Claude desktop app, use **Claude Code**: choose to work on
   **this computer** and pick the folder you just made.
3. **Paste this as your first message:**

   > Set up the Outlook connector: clone https://github.com/jfeuerman24/outlook-desktop-mcp.git
   > into this folder, then read its CLAUDE.md and walk me through it.

4. Claude explains what the tool does, checks the PC, installs it, and adds it to Claude's own
   settings (keeping a backup of the old settings file). It then asks you to **quit Claude from
   the system tray and open it again** — that is the only thing you do yourself. Come back to
   the folder and say **"test it"**.

The first time Claude touches Outlook, Outlook asks whether to allow it. Tick **Allow access
for**, pick **10 minutes**, and carry on; see [the security prompt](#outlook-programmatic-access-security-prompts)
for the permanent options.

Already downloaded the ZIP from GitHub? Unzip it, open that folder with Claude Code, and say
**"let's get started"**. Prefer to do it by hand, or want it in Claude Code's terminal? See
[SETUP.md](SETUP.md). You can also double-click `install.cmd` in the folder, which does the
install and the settings change without Claude.

### Uninstall

Ask Claude to remove it, or run `.venv\Scripts\python.exe scripts\configure_claude_desktop.py --remove`,
restart Claude, and delete the folder.

## Requirements

- **Outlook Desktop (Classic)** — the `OUTLOOK.EXE` that comes with Microsoft 365 / Office. The new "modern" Outlook (`olk.exe`) does **not** support COM
- **Windows 10 or 11**, x64 or ARM64
- **Python 3.12+**
- **Outlook must be running** when the MCP server starts

On ARM64, all dependencies (`pywin32`, `mcp`, `pydantic-core`, `cryptography`, `cffi`, `rpds-py`) have prebuilt `win_arm64` wheels — see the [ARM64 install notes](#windows-arm64) below for the one extra `pip` flag you need.

### Outlook "Programmatic Access" security prompts

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

## Tools

### Email

| Tool | Description |
|------|-------------|
| `list_emails` | List recent emails from any folder, with optional unread, date, sender and subfolder filters |
| `read_email` | Read full email content by entry ID or subject search (quoted history trimmed unless `include_history`) |
| `search_emails` | Full-text search across email subjects and bodies, optionally by sender and across subfolders |
| `get_conversation` | Every message of a thread, oldest first, each once, across folders and dates |
| `list_folders` | Browse the folder hierarchy with item counts |
| `draft_email` | Write a new email to Drafts with To/CC/BCC and open it in Outlook. **Not sent.** |
| `draft_reply` | Write a reply or reply-all above the quoted thread, save to Drafts, open in Outlook. **Not sent.** |
| `mark_as_read` | Mark a specific email as read |
| `mark_as_unread` | Mark a specific email as unread |
| `move_email` | Move an email to Archive, Trash, or any folder |

### Calendar

| Tool | Description |
|------|-------------|
| `list_calendars` | List every calendar folder in an account (Birthdays, holidays, shared, subscribed) |
| `list_events` | List events within a date range, from one calendar or `calendar="all"` |
| `get_event` | Read full event details by entry ID |
| `search_events` | Search calendar events by keyword within a date range |
| `create_event` | Create a personal calendar appointment (no attendees) |
| `draft_meeting` | Put a meeting with attendees on the calendar and open it in Outlook. **Invitations not sent.** |
| `update_event` | Change an event's subject, time, location, etc. Saved, not sent |
| `delete_event` | Delete an appointment or a received meeting. Refuses meetings you organised |

### Tasks

| Tool | Description |
|------|-------------|
| `list_tasks` | List pending or completed tasks, sorted by due date |
| `get_task` | Read full task details including body and completion status |
| `create_task` | Create a new task with subject, due date, importance |
| `complete_task` | Mark a task as complete |
| `delete_task` | Remove a task |

### Attachments, accounts

| Tool | Description |
|------|-------------|
| `list_accounts` | List the mailboxes in the Outlook profile |
| `list_attachments` | List all attachments on an email or calendar event |
| `save_attachment` | Save an attachment to a local directory (writes to disk, never to the mailbox) |

### Categories, Rules, Out of Office

| Tool | Description |
|------|-------------|
| `list_categories` | List all available color categories in Outlook |
| `set_category` | Set or clear categories on any email, event, or task |
| `list_rules` | List all mail rules with enabled/disabled status |
| `toggle_rule` | Enable or disable a mail rule by name |
| `get_out_of_office` | Check whether Out of Office auto-reply is on or off |

**Total: 31 tools.**

### Nothing sends

The three tools that would normally send instead save and hand over to you:

- `draft_email` writes a new message to the Drafts folder of the chosen account.
- `draft_reply` builds Outlook's own reply or reply-all (recipients, quoted thread, signature) and puts the text above it.
- `draft_meeting` puts the meeting on your calendar with the attendees filled in and the invitations unsent.

Each one saves the item and, unless called with `open_in_outlook=false`, opens it in an Outlook
window so you can read it and press **Send**. The tool result says `"status": "draft_saved"`
and `"sent": false`, with the entry ID and the folder, so Claude can tell you where it is.
Drafts get your normal signature and the font described below.

Deliberately missing: there is no tool to send, to accept or decline an invitation, or to
cancel a meeting you organised, because every one of those sends a message. Do those in Outlook.

### Working with several calendars

```
list_calendars account="personal"
→ Calendar (default), Calendar/Birthdays, Calendar/United States holidays, Calendar/Family

list_events start_date="2026-10-27" end_date="2026-10-29" calendar="all" account="personal"
→ events from all four, each tagged with its calendar, sorted by start
```

To check availability across accounts, call `list_events` with `calendar="all"` once per account and
compare the results.

### Draft font

Drafts are written as HTML in Aptos 12 pt black so they look the same everywhere. Override with
environment variables before the server starts: `OUTLOOK_MCP_FONT_FAMILY`, `OUTLOOK_MCP_FONT_SIZE_PT`,
`OUTLOOK_MCP_FONT_COLOR`, or `OUTLOOK_MCP_PLAIN_TEXT=1` for plain text.

## How it works

Windows Outlook (Classic) exposes the Outlook Object Model over COM — the same automation interface
add-ins and VBA have used for 20 years. All COM calls run on one dedicated thread using the
Single-Threaded Apartment model, as COM requires. The async MCP event loop dispatches each tool call to
that thread through a queue and awaits the result, so COM threading rules are respected and the MCP
protocol stays non-blocking.

```
MCP tool call (async)
  → bridge.call(func, args)
    → queued to STA thread
      → func(outlook, namespace, args) executes on COM thread
    → result returned via threading.Event
  → JSON response back to MCP client
```

Each tool's inner function receives the live `Outlook.Application` and `MAPI.Namespace` COM objects and
works directly with the Outlook Object Model — `GetItemFromID`, `Items.Add`, `Items.Restrict` with DASL
filters, and so on. See [`com_bridge.py`](src/outlook_desktop_mcp/com_bridge.py).

## Install from Source

`install.ps1` performs everything in this section (and the Claude registration) automatically; the steps are listed for reference and for troubleshooting.

### Windows (x64)

```bash
git clone https://github.com/jfeuerman24/outlook-desktop-mcp.git
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
git clone https://github.com/jfeuerman24/outlook-desktop-mcp.git
cd outlook-desktop-mcp
& "C:\Program Files\Python312-arm64\python.exe" -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install --only-binary=:all: pywin32 mcp
python -m pip install --no-deps -e .
python .venv\Scripts\pywin32_postinstall.py -install
```

The base `mcp` package is sufficient for running the stdio server — the `[cli]` extra is only needed for the `mcp` developer CLI tools (`mcp dev`, `mcp inspector`), which aren't used at runtime.

Register from source the same way as x64.

## Usage Examples

Once registered, just talk to Claude naturally:

- *"Show me my 10 most recent inbox emails"*
- *"Read the email from Taylor about MLADS and give me the whole thread"*
- *"What's on all my calendars this week?"*
- *"Save the attachment from that email to my Downloads folder"*
- *"Draft a reply to Taylor saying I can do Thursday"* — the reply opens in Outlook; you press Send
- *"Draft an email to alice@example.com about the project update"*
- *"Set up a 30-minute meeting with bob@example.com tomorrow at 2pm"* — the invitation opens in Outlook; you press Send
- *"Create a task to review the quarterly report, due Friday, high importance"*
- *"Mark that email as read and move it to archive"*
- *"What categories do I have? Set this email to 'Follow-up'"*

## Why Not Microsoft Graph?

| | Microsoft Graph | outlook-desktop-mcp |
|---|---|---|
| Entra app registration | Required | Not needed |
| Admin consent | Required for mail permissions | Not needed |
| OAuth token management | You handle refresh tokens | Not needed |
| Tenant configuration | Required | Not needed |
| Works offline / cached | No | Yes (reads from local cache) |
| Can send on your behalf | Yes | Never — drafts only |
| Setup time | 30-60 minutes | 2 minutes |
| Auth requirement | **Your own OAuth flow** | **Outlook is open** |

## Project Structure

```
outlook-desktop-mcp/
  src/outlook_desktop_mcp/
    entrypoint.py            # Windows check → server
    server.py                # The MCP server (31 tools, COM automation, drafts only)
    com_bridge.py            # Async-to-COM threading bridge
    tools/
      _folder_constants.py   # Outlook enums and constants
      calendars.py           # Every calendar folder in a store; calendar argument resolution
      conversation.py        # Conversation tree walk + ConversationTopic fallback
    utils/
      formatting.py          # Email/event/task data extraction
      text.py                # Quoted-history trimming, external-sender banner removal (pure)
      dates.py               # pywintypes → ISO 8601 with offset (pure)
      addresses.py           # Exchange DN → SMTP resolution, recipients
      html.py                # Draft font styling, injecting text above signature/quoted history (pure)
      errors.py              # COM error formatting
  tests/
    test_text.py             # unit: trimming (no Outlook)
    test_dates.py            # unit: ISO dates (no Outlook)
    test_html.py             # unit: draft HTML (no Outlook)
    live_check.py            # live: the fork's additions over stdio, reads only, changes nothing
    phase1_com_test.py       # Email COM validation (writes a draft, sends nothing)
    phase3_mcp_test.py       # Email MCP test (writes a draft, sends nothing)
    calendar_com_test.py     # Calendar COM validation (prepares an unsent meeting)
    calendar_mcp_test.py     # Calendar MCP test (prepares an unsent meeting)
    extras_com_test.py       # Tasks/attachments/categories/rules/OOF COM test
    extras_mcp_test.py       # Tasks/attachments/categories/rules/OOF MCP test
  CLAUDE.md                 # Guide Claude follows to install and explain the tool for a new user
  install.ps1               # Checks the PC, installs Python if needed, builds .venv, registers with Claude
  install.cmd               # Double-click wrapper for install.ps1
  scripts/
    configure_claude_desktop.py  # Adds/removes the server in claude_desktop_config.json, with backup
  outlook-desktop-mcp.cmd   # Windows launcher script
  pyproject.toml
```

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for the branching strategy and development setup.

## License

See [LICENSE](LICENSE) file.
