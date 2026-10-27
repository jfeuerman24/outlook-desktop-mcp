# outlook-desktop-mcp

[![PyPI](https://img.shields.io/pypi/v/outlook-desktop-mcp)](https://pypi.org/project/outlook-desktop-mcp/)
[![Python](https://img.shields.io/pypi/pyversions/outlook-desktop-mcp)](https://pypi.org/project/outlook-desktop-mcp/)
[![Platform](https://img.shields.io/badge/platform-Windows-blue)]()

**Lets Claude use the Outlook that is already open on your Windows PC.**

Claude can look through your mail, calendars and tasks, find things, pull a whole thread
together, and write emails, replies and meeting invitations for you. **It never sends
anything.** Whatever it writes goes to your Drafts folder, or onto your calendar with the
invitations unsent, for you to read and send yourself. Unsent meetings open in an Outlook
window so you can find them; email drafts wait quietly in Drafts.

There is no Microsoft account setup, nothing for an admin to approve, and no password to hand
over. The server talks to the Outlook program on your computer the way an add-in would, so it
has the access you already have, and only while Outlook is open.

| Claude can | Claude cannot |
|---|---|
| List, search and read mail, calendars and tasks. Save attachments to a folder. Write draft emails and replies. Prepare meeting invitations. Create and change appointments and tasks. Move and mark mail, set categories, switch rules on and off. | Send anything. Accept or decline invitations. Cancel meetings you organised. |

> **This is a fork** of [Aanerud/outlook-desktop-mcp](https://github.com/Aanerud/outlook-desktop-mcp)
> (AGPL-3.0, see [LICENSE](LICENSE)). It is Windows-only and never sends; the changes are listed
> under [What this fork changes](#what-this-fork-changes).

## Getting started

### What you need

- A Windows PC with **classic Outlook** signed in and open. The new Outlook app and Outlook on
  the web do not work with this.
- The **Claude desktop app**. The easiest setup is done by **Claude Code** inside the app, which
  needs a paid plan (Pro, Max, Team or Enterprise). There is a path without Claude further down.
- About five minutes. You will not type any commands; Claude runs them and explains as it goes.

### Set up Outlook first

Claude sees exactly the mailboxes that are in Outlook on this PC, so add them before you start:

1. Open classic Outlook and make sure the **New Outlook** switch in its top-right corner is off.
2. Add every mailbox you want Claude to use: **File → Add Account**, sign in, repeat.
   Microsoft 365, Exchange, Outlook.com and Gmail or other IMAP accounts all work; each one
   becomes an account Claude can name.
3. Shared mailboxes you have full access to appear on their own under your work account. If one
   is missing: **File → Account Settings → Account Settings**, select the account, then
   **Change → More Settings → Advanced → Add**.
4. Leave Outlook open. The connector only works while it is running.

During setup Claude lists the accounts it can see. If one you expected is not there, it tells
you how to add it.

### Easiest: let Claude do it

1. **Make an empty folder** for it, for example `C:\Users\<you>\outlook-desktop-mcp`
   (open your user folder in File Explorer, right-click → **New** → **Folder**). Avoid OneDrive folders.
2. **Open it in Claude.** In the Claude desktop app, use **Claude Code**: choose to work on
   **this computer** and pick the folder you just made.
3. **Paste this as your first message:**

   > Set up the Outlook connector: clone https://github.com/jfeuerman24/outlook-desktop-mcp.git
   > into this folder, then read its CLAUDE.md and walk me through it.

4. Claude explains what the tool does, checks the PC, installs it, and adds it to Claude's own
   settings (keeping a backup of the old settings file). It then asks you to **quit Claude from
   the system tray and open it again**. That is the only thing you do yourself. Come back to
   the folder and say **"test it"**.

The first time Claude touches Outlook, Outlook asks whether to allow it. Tick **Allow access
for**, pick **10 minutes**, and carry on; see [the security prompt](#outlook-programmatic-access-security-prompts)
for the permanent options.

**Updating** works the same way: open the folder in Claude Code and say **"update it"**. Claude
pulls the latest version, runs the installer again and asks you to restart Claude.

Already downloaded the ZIP from GitHub? Unzip it, open that folder with Claude Code, and say
**"let's get started"**.

### Without Claude

No Claude subscription is needed for this path; the server registers itself with whichever
Claude desktop app is installed.

- **Install:** download the repository as a ZIP from GitHub (or clone it), unzip it into a folder
  outside OneDrive, and double-click `install.cmd`. It checks the PC, installs the server and
  adds it to the Claude desktop app's settings. Then quit Claude from the system tray, make sure
  Outlook is open, and start Claude again.
- **Update:** pull the latest version with VS Code's Source Control view, GitHub Desktop or
  `git pull` (if you unzipped, download the ZIP again and replace the files, keeping `.venv`),
  then double-click `install.cmd` again and restart Claude.

Prefer to do each step by hand, or want it in Claude Code's terminal? See [SETUP.md](SETUP.md).

### Uninstall

Ask Claude to remove it, or run `.venv\Scripts\python.exe scripts\configure_claude_desktop.py --remove`,
restart Claude, and delete the folder.

## What this fork changes

| Change | Where |
|---|---|
| **Never sends.** Upstream's `send_email`, `reply_email` and `create_meeting` are replaced by `draft_email`, `draft_reply`, `draft_forward` and `draft_meeting`, which save the item for you to send. `respond_to_meeting` is gone and `delete_event` refuses meetings you organised, because both would send a message. | `server.py` |
| **Windows only.** The macOS/AppleScript server was removed. | `entrypoint.py` |
| **Claude installs and explains it.** `CLAUDE.md` is the guide Claude follows; `install.ps1` / `install.cmd` do the install and register the server with the desktop app, including the Microsoft Store build. | `CLAUDE.md`, `install.ps1`, `scripts/` |
| **Drafts are written in the chat first.** The draft tools tell Claude to show the full draft and ask before putting it in Outlook; drafts are single-spaced in Aptos 12 pt with your signature. Email drafts are saved quietly to Drafts; unsent meetings open in an Outlook window (the `open_in_outlook` setting). | `server.py`, `utils/html.py` |
| **Thread context before a reply.** `reply_context` returns the recipients Outlook's own Reply would address, who is on the thread and a digest of it, so the chat preview starts with context. Every message carries `thread_depth`. `open_in_outlook` opens any item in Outlook on request. | `server.py`, `tools/conversation.py` |
| **Settings by conversation.** `set_setting` / `get_settings` for the default account, hidden accounts, draft window and font, stored outside the repo. | `utils/settings.py` |
| **Attachments read in place, files attached from disk.** `read_attachment` (PDF, Word, Excel, PowerPoint, CSV, text, .msg), `add_attachment`, and an `attachments` argument on every draft tool. | `utils/extract.py`, `server.py` |
| **Fast searches.** Outlook does the filtering (sender, recipient, subject, date) instead of Python walking messages, so a sender search of an 11,000-message inbox takes a second and sees every message. Shared mailboxes that are not cached locally are skipped by `account="all"`. Long searches return partial results before the time limit instead of an error. `legacy_mode` runs the old paths. | `utils/dasl.py`, `utils/deadline.py`, `com_bridge.py`, `server.py` |
| **One-call lookups.** `contact_history` (one person across every account and calendar), `account="all"` on `list_emails`/`search_emails`, `update_draft`, `health_check`. | `server.py` |
| **Every calendar, not just the default.** `list_calendars`; `list_events` and `search_events` take `calendar` (`""`, `"all"`, or a name/path). | `tools/calendars.py` |
| **Thread reconstruction.** `get_conversation` walks Outlook's conversation across folders and dates, deduped by Message-ID, oldest first. | `tools/conversation.py` |
| **Quoted history trimmed, Exchange addresses resolved, ISO dates.** `read_email` returns the message's own text; senders inside your own organisation come back as real addresses; dates carry the local UTC offset. | `utils/text.py`, `utils/addresses.py`, `utils/dates.py` |
| **Offline unit tests** (`python -m unittest discover -s tests`) and a GitHub Actions workflow that runs them. | `tests/test_*.py` |

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
| `list_emails` | List recent emails from any folder, with optional unread, date, sender and subfolder filters; `account="all"` for every mailbox |
| `read_email` | Read full email content by entry ID or subject search (quoted history trimmed unless `include_history`); `thread_depth` says how deep in a thread it sits |
| `search_emails` | Find mail by sender (name, address or domain), subject text and date, in about a second even in a big folder; `include_body` also scans message text (slow); `account="all"` for every mailbox |
| `get_conversation` | Every message of a thread, oldest first, each once, across folders and dates; `max_body_chars` for a cheap look at long threads |
| `reply_context` | What a reply preview needs: the To/CC Outlook's own Reply or Reply All would use, who is on the thread, a digest of the earlier messages, the message in full. Creates nothing |
| `open_in_outlook` | Open an email, draft, meeting, appointment or task in its own Outlook window when you ask. Changes nothing |
| `contact_history` | One known person across every account: last received, last sent, recent messages both ways, last and next meeting. The broadest search; use after a targeted one |
| `list_folders` | Browse the folder hierarchy with item counts |
| `draft_email` | Write a new email to Drafts with To/CC/BCC and optional file attachments. **Not sent.** |
| `draft_reply` | Write a reply or reply-all above the quoted thread, save to Drafts, optional attachments. **Not sent.** |
| `draft_forward` | Forward an email with Outlook's own Forward() (FW: subject, header block, quoted message, attachments), your note above it, saved to Drafts. Can forward the newest message of the thread and pull in attachments from earlier messages. **Not sent.** |
| `update_draft` | Change an existing draft's note, subject or recipients instead of making a new one. **Not sent.** |
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
| `draft_meeting` | Put a meeting with attendees on the calendar; it opens in Outlook for you to check. **Invitations not sent.** |
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
| `list_accounts` | List the mailboxes in the Outlook profile (hidden ones left out unless asked); flags shared mailboxes that all-account searches skip |
| `list_attachments` | List all attachments on an email or calendar event |
| `save_attachment` | Save an attachment to a local directory (writes to disk, never to the mailbox) |
| `read_attachment` | Read the text of a PDF, Word, Excel, PowerPoint, CSV, text or .msg attachment without saving it |
| `add_attachment` | Attach files from this PC to an existing draft |

### Categories, Rules, Out of Office

| Tool | Description |
|------|-------------|
| `list_categories` | List all available color categories in Outlook |
| `set_category` | Set or clear categories on any email, event, or task |
| `list_rules` | List all mail rules with enabled/disabled status |
| `toggle_rule` | Enable or disable a mail rule by name |
| `get_out_of_office` | Check whether Out of Office auto-reply is on or off |

### Settings, health

| Tool | Description |
|------|-------------|
| `get_settings` | Show every setting, its value and where it came from |
| `set_setting` | Change a setting by asking Claude (default account, hidden accounts, draft window, font); takes effect at once |
| `health_check` | Server version, Outlook running, accounts loaded, settings, and whether Outlook's permission prompt is blocking |

**Total: 41 tools.**

### Nothing sends

The four tools that would normally send instead save and hand over to you:

- `draft_email` writes a new message to the Drafts folder of the chosen account.
- `draft_reply` builds Outlook's own reply or reply-all (recipients, quoted thread, signature) and puts the text above it.
- `draft_forward` builds Outlook's own forward (FW: subject, header block, quoted message, attachments) and puts your note above it.
- `draft_meeting` puts the meeting on your calendar with the attendees filled in and the invitations unsent.

Each one saves the item so you can open it, read it and press **Send**. By default an unsent
meeting opens in an Outlook window at once, so it is not lost on the calendar, and an email draft
waits in Drafts without a window; the `open_in_outlook` setting (`none`, `meetings`, `all`)
changes that, and asking Claude to "open it in Outlook" opens any item. The tool result says
`"status": "draft_saved"` and `"sent": false`, with the entry ID and the folder, so Claude can
tell you where it is. Drafts get your normal signature and the font described below.

Deliberately missing: there is no tool to send, to accept or decline an invitation, or to
cancel a meeting you organised, because every one of those sends a message. Do those in Outlook.

### Context before a reply

You will often ask for a reply to an email you have not opened. Before showing a reply or
forward draft, Claude calls `reply_context`, which returns the recipients Outlook's own Reply or
Reply All would address, who is on the thread, a short digest of the earlier messages and the
message being answered in full, without creating anything. The draft preview in the chat then
starts with a few lines of context. Every message also carries `thread_depth` (0 for the first
message of a thread, 1 for a reply to it, and so on), so Claude can tell when there is history
it has not read. "Open that email in Outlook" (`open_in_outlook`) shows it in Outlook's own
window, thread and all.

### Searching quickly

Outlook filters by sender, recipient, subject and date itself, in well under a second even in
a folder of 10,000 messages, so `search_emails` with a `sender` (name, address or domain) or a
subject word is the fast way to find a thread, and Claude is told to start there. Three things
are slower and are handled explicitly:

- **Body text.** `include_body=true` opens every message in the folder, 10 to 30 seconds on a
  big one. On a PC where Windows Search has indexed Outlook's message bodies the server detects
  that once per session and uses the index instead; `health_check` says which applies.
- **Shared mailboxes that are not cached locally.** Every folder listing in one is a server
  round trip (18 to 30 seconds just to find its calendars). `account="all"` skips them and lists
  them under `accounts_skipped`; name the mailbox to search it, or set `include_shared_in_all`.
- **Time limits.** A search that cannot finish within `call_timeout_s` (default 60) stops a few
  seconds early and returns what it found with `"partial": true` and a note, rather than an
  error. A call made while an earlier one has overrun fails at once with "Outlook is busy".

If a new search path misbehaves, `set_setting("legacy_mode", "true")` runs the pre-0.6 code
immediately, no restart. `tests/perf_check.py` times the searches that used to be slow (read-only).

### Working with several calendars

```
list_calendars account="personal"
→ Calendar (default), Calendar/Birthdays, Calendar/United States holidays, Calendar/Family

list_events start_date="2026-10-27" end_date="2026-10-29" calendar="all" account="personal"
→ events from all four, each tagged with its calendar, sorted by start
```

To check availability across accounts, call `list_events` with `calendar="all"` once per account and
compare the results.

### Settings

Settings are changed by asking Claude ("make my HGP mailbox the default account", "hide the
Pickleball mailbox"), which calls `set_setting`. They are stored in
`%LOCALAPPDATA%\outlook-desktop-mcp\settings.json`, apply immediately, and survive restarts
and updates. `get_settings` shows them. Each one can also be set as an environment variable in
Claude's config, which then overrides the file.

| Setting | Env var | Default | What it does |
|---|---|---|---|
| `default_account` | `OUTLOOK_MCP_DEFAULT_ACCOUNT` | Outlook's default mailbox | Account used when a request does not name one |
| `hide_accounts` | `OUTLOOK_MCP_HIDE_ACCOUNTS` | none | Mailboxes (name substrings, `;`-separated) left out of `list_accounts` and `account="all"`; still reachable by name |
| `open_in_outlook` | `OUTLOOK_MCP_OPEN_IN_OUTLOOK` | `meetings` | Which saved drafts open in an Outlook window: `meetings` (unsent invitations only), `all`, or `none`. `open_drafts` (true/false) is the older name and still works until this one is set |
| `font_family`, `font_size_pt`, `font_color` | `OUTLOOK_MCP_FONT_*` | Aptos, 12, black | Font for the text the server writes into drafts |
| `plain_text` | `OUTLOOK_MCP_PLAIN_TEXT` | off | Write drafts as plain text instead of styled HTML |
| `include_shared_in_all` | `OUTLOOK_MCP_INCLUDE_SHARED_IN_ALL` | off | Include shared, non-cached mailboxes in `account="all"` searches (slow) |
| `call_timeout_s` | `OUTLOOK_MCP_CALL_TIMEOUT_S` | 60 | Seconds a tool call may take; long searches return partial results just before it |
| `legacy_mode` | `OUTLOOK_MCP_LEGACY_MODE` | off | Run the pre-0.6 search paths (slow); for troubleshooting |

### Reading attachments

`read_attachment` pulls the text out of PDF, Word (.docx), Excel (.xlsx), PowerPoint (.pptx),
CSV, plain-text, HTML and attached-message (.msg) files so Claude can answer "what does the lease
say about assignment" without anyone opening the file. The attachment is copied to a temporary
folder, read, and deleted; nothing in the mailbox changes. Long files are capped (default 20,000
characters) and the result says so; ask for a page range (`pages="1-3"`) or one sheet
(`sheet="Summary"`) instead of raising the cap. Scanned PDFs have no text layer and are reported
as such; old `.doc`/`.xls`/`.ppt` files are not supported.

### Checking that it works

`health_check` reports the server version, whether Outlook is reachable, which accounts are
loaded and which is the default, which are shared mailboxes skipped by all-account searches,
whether body-text search can use the Windows Search index, the current settings, and whether
Outlook is blocking address access behind its permission prompt. It is the first thing to run
when something fails, particularly from a phone where the laptop screen cannot be seen.

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
- *"When did we last hear from Chad, and do we have anything on the calendar with him?"*
- *"What does the lease in Mitch's email say about assignment?"* — read straight from the PDF attachment
- *"Make my HGP mailbox the default account"*
- *"Read the email from Taylor about MLADS and give me the whole thread"*
- *"What's on all my calendars this week?"*
- *"Save the attachment from that email to my Downloads folder"*
- *"Draft a reply to Taylor saying I can do Thursday"* — Claude first shows a few lines about the thread and the recipients Outlook would use, then the reply; it is saved to Drafts and you press Send
- *"Change that draft to say Friday instead"* — the same draft is updated, not a second one
- *"Open that email in Outlook"* — the message opens in Outlook's own window
- *"Is the Outlook connector working?"* — `health_check`
- *"Draft an email to alice@example.com about the project update"*
- *"Set up a 30-minute meeting with bob@example.com tomorrow at 2pm"* — the invitation is on your calendar, unsent; you press Send
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
    server.py                # The MCP server (41 tools, COM automation, drafts only)
    com_bridge.py            # Async-to-COM threading bridge
    tools/
      _folder_constants.py   # Outlook enums and constants
      calendars.py           # Every calendar folder in a store; calendar argument resolution
      conversation.py        # Conversation tree walk + ConversationTopic fallback; thread depth
    utils/
      formatting.py          # Email/event/task data extraction
      text.py                # Quoted-history trimming, external-sender banner removal (pure)
      dates.py               # pywintypes → ISO 8601 with offset (pure)
      addresses.py           # Exchange DN → SMTP resolution, recipients
      html.py                # Draft font styling, note markers, injecting text above signature/quoted history (pure)
      settings.py            # Settings file + env var resolution (pure)
      extract.py             # Text out of PDF/Word/Excel/PowerPoint/CSV attachments (pure)
      dasl.py                # Outlook Restrict filter builders: sender, recipient, subject, body, dates (pure)
      deadline.py            # Soft deadlines so long searches return partial results in time (pure)
      errors.py              # COM error formatting
  tests/
    test_text.py             # unit: trimming (no Outlook)
    test_dates.py            # unit: ISO dates (no Outlook)
    test_html.py             # unit: draft HTML (no Outlook)
    test_settings.py         # unit: settings resolution (no Outlook)
    test_extract.py          # unit: attachment text extraction (no Outlook)
    test_dasl.py             # unit: filter builders (no Outlook)
    test_deadline.py         # unit: deadlines and bridge request bookkeeping (no Outlook)
    test_calendars.py        # unit: calendar discovery cache with a fake store (no Outlook)
    test_conversation.py     # unit: thread depth from ConversationIndex (no Outlook)
    live_check.py            # live: the fork's additions over stdio, reads only, changes nothing
    perf_check.py            # live: times the searches that used to be slow, reads only
    phase1_com_test.py       # Email COM validation (writes a draft, sends nothing)
    phase3_mcp_test.py       # Email MCP test (writes a draft, sends nothing)
    calendar_com_test.py     # Calendar COM validation (prepares an unsent meeting)
    calendar_mcp_test.py     # Calendar MCP test (prepares an unsent meeting)
    extras_com_test.py       # Tasks/attachments/categories/rules/OOF COM test
    extras_mcp_test.py       # Tasks/attachments/categories/rules/OOF MCP test
  CLAUDE.md                 # Guide Claude follows to install and explain the tool for a new user
  install.ps1               # Checks the PC, installs Python and Git if needed, builds .venv, registers with Claude
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
