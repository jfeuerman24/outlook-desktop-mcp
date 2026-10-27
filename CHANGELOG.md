# Changelog

All notable changes to this fork. The version is in `pyproject.toml` and
`outlook_desktop_mcp.__version__`; `outlook-desktop-mcp.exe --version` and `health_check` report it.

## Unreleased

### Fixed
- `read_attachment` on an Excel file left the workbook open when the requested sheet did
  not exist (or reading raised), so on Windows the temporary copy could not be deleted. The
  workbook is now closed on every path. This is what failed the unit tests in CI for 0.5.0.

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
