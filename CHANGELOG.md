# Changelog

All notable changes to this fork. The version is in `pyproject.toml` and
`outlook_desktop_mcp.__version__`; `outlook-desktop-mcp.exe --version` and `health_check` report it.

## 0.7.0 - 2026-10-27

### Added
- **`reply_context`**: before a reply or forward draft is shown in the chat, one call returns
  the recipients Outlook's own Reply or Reply All would address, who is on the thread, a digest
  of the earlier messages and the message being answered in full. Nothing is created or saved.
  The server instructions and the draft tools now ask for a short Context block above every
  reply preview, because the user often has not opened the email in Outlook.
- **`thread_depth`** on every email record: 0 for the first message of a thread, 1 for a reply
  to it, and so on, read from Outlook's ConversationIndex at no cost.
- **`open_in_outlook`** tool: opens any item (email, draft, unsent meeting, appointment, task)
  in an Outlook window when the user asks. Changes nothing.
- `install.ps1` looks for Git and installs it with winget when missing, so the folder can be
  updated with `git pull` later. CLAUDE.md has an "update it" procedure; the README says how to
  update with or without Claude.

### Changed
- **Meetings open, emails do not.** The `open_drafts` boolean is replaced by `open_in_outlook`
  with `none`, `meetings` (default) or `all`. An unsent meeting opens in an Outlook window so it
  is not lost on the calendar; an email draft waits in Drafts. `open_drafts` still works as the
  older name until `open_in_outlook` is set. The per-call `open_in_outlook` argument is gone
  from the draft tools (Claude kept setting it); the tool of the same name covers "open it for me".
- README: what you need (classic Outlook with your accounts added; the Claude desktop app, with
  a paid plan for the guided setup), how to add mailboxes to Outlook first, a path without
  Claude, and how to update. SETUP.md no longer claims free plans lack local servers.
  CONTRIBUTING.md describes this fork rather than upstream's release flow.
- The server instructions and the mail, calendar and task entry-point tools now say this is
  the user's primary mailbox and calendar, so Claude Desktop reaches for Outlook rather than a
  Gmail or Google Calendar connector when the request does not name a service.

## 0.6.0 - 2026-10-27

Searching is now fast. Measured on a profile with five accounts, one of them an 11,000-message
inbox: the nine calls behind two real sessions took 250 s with three timeouts before, 26 s
with none after.

### Changed
- **Outlook filters, Python no longer walks.** The `sender` filter on `list_emails` and
  `search_emails` is a native filter on the sender's name, directory address and SMTP
  address (1 s on 11,000 messages, and complete: the old scan only looked at the newest
  1,000 messages per folder and silently missed older mail from that sender).
- **`search_emails` matches subjects by default.** `include_body=true` also scans message
  text, which opens every message and is slow on big folders; use it last and with a date
  range. Where Windows Search has indexed Outlook bodies, the server detects it once per
  session and body search is fast too (`health_check` reports which).
- **`contact_history`** skips plain appointments before reading anything, reads three
  fields per meeting instead of formatting every occurrence, lets Outlook match recipients
  by the names the person is known by (found from the received side, so an address or
  domain works too), skips Birthdays and holiday calendars, and caches each account's
  calendar list for ten minutes.
- **Shared mailboxes that are not cached locally are skipped by `account="all"`** and
  listed under `accounts_skipped`; every folder listing in such a mailbox is a server round
  trip (18 to 30 s just to find its calendars). Name the mailbox to search it, or set
  `include_shared_in_all`. `list_accounts` and `health_check` flag them.
- **Time limits return results instead of errors.** Long searches stop a few seconds
  before the limit (`call_timeout_s`, default 60) and return what they found with
  `"partial": true`. A call queued behind an overrunning call fails at once with a clear
  message instead of also timing out; a queued call whose caller has given up is skipped.
- Tool descriptions and the server instructions steer Claude to sender and subject
  searches in the default account first, and to `contact_history` only for one known
  person.
- Draft tool descriptions no longer claim drafts open in a window by default (they have
  not since 0.5.0).

### Added
- Settings `include_shared_in_all`, `call_timeout_s` and `legacy_mode`. Legacy mode runs
  the pre-0.6 search paths unchanged, as a runtime fallback; it needs no restart.
- `tests/perf_check.py`: read-only timing of the searches above, for before/after numbers.
- Unit tests for the filter builders, deadlines, the bridge's request bookkeeping and the
  calendar cache.

### Fixed
- `read_attachment` on an Excel file left the workbook open when the requested sheet did
  not exist (or reading raised), so on Windows the temporary copy could not be deleted. The
  workbook is now closed on every path. This is what failed the unit tests in CI for 0.5.0.
- Package metadata pointed at the upstream repository.

## 0.5.0 - 2026-10-27

### Added
- **Settings by conversation.** `get_settings` and `set_setting` tools; settings live in
  `%LOCALAPPDATA%\outlook-desktop-mcp\settings.json`, apply immediately, and survive restarts
  and `git pull`. Environment variables still override.
  - `default_account`: the mailbox tools use when a request does not name one.
  - `hide_accounts`: mailboxes left out of `list_accounts` and `account="all"`.
  - `open_drafts`, `font_family`, `font_size_pt`, `font_color`, `plain_text` moved here.
- **`read_attachment`**: text of PDF, Word, Excel, PowerPoint, CSV, text, HTML and attached
  `.msg` files without saving or opening them, with page/sheet selection and a size cap.
  New libraries: pypdf, python-docx, openpyxl, python-pptx (all pure Python; `install.ps1`
  adds them to an existing install).
- **`contact_history`**: one person across every account, both directions, plus last and
  next meeting with them.
- **`update_draft`**: change the note, subject or recipients of an existing draft. The note
  the server writes is now wrapped in a marker so it can be replaced without touching the
  signature or quoted thread.
- **`add_attachment`** and an `attachments` argument on `draft_email`, `draft_reply` and
  `draft_forward` to attach files from this PC.
- **`draft_forward`**: Outlook's own Forward() with the note above; can forward the newest
  message of the thread and pull attachments from earlier messages.
- **`health_check`**: version, Outlook reachable, accounts, settings, and a timed probe that
  reports when Outlook's "A program is trying to access…" prompt is blocking.
- `account="all"` on `list_emails` and `search_emails`; results carry their account.
- `max_body_chars` on `get_conversation` for a cheap look at long threads.
- `--version` on the server executable.
- Offline unit tests for settings, extraction and draft note markers.

### Changed
- Draft tools no longer open a window by default; they save to Drafts (or the calendar) and
  say where. `open_drafts` turns the window back on.
- Draft text is single-spaced with one blank line between blocks, and plain-text drafts
  render numbered/bulleted lists, `> ` indented lines and `**bold**`.
- The draft tools' descriptions tell Claude to show the full draft in the chat and ask
  before putting it in Outlook.
- `list_accounts` marks the account tools use by default and leaves hidden accounts out.

### Fixed
- Senders inside the user's own Exchange organisation came back as
  `/O=EXCHANGELABS/…`; the message's own SMTP address property is now read first.

## 0.4.0

Fork of Aanerud/outlook-desktop-mcp: drafts only (no tool can send), Windows only, every
calendar, conversation reconstruction, sender/subfolder filters, SMTP address resolution,
ISO dates, trimmed quoted history, Aptos draft styling, Claude-guided installer
(`CLAUDE.md`, `install.ps1`, `install.cmd`, `scripts/configure_claude_desktop.py`).
