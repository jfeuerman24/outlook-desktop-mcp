# Changelog

All notable changes to this fork. The version is in `pyproject.toml` and
`outlook_desktop_mcp.__version__`; `outlook-desktop-mcp.exe --version` and `health_check` report it.

## 0.11.0 - 2026-10-29

### Added
- **`set_flag`**: flag an email for follow-up (no date, today, tomorrow, this week, next
  week, or a date), mark the flag complete, or clear it. Outlook refuses to flag drafts.
- **Inline notes on forwards.** `draft_forward` takes `inline_notes`, the same
  `{"after", "text"}` list as `draft_reply`'s `inline_replies`, for "forward this to Chris
  with my comments under each point". `update_draft` revises them.
- **Attendee responses.** `get_event` lists every attendee with their response and groups
  them (`responses`: accepted, tentative, declined, no_response). Outlook tracks responses
  only for the organiser; `responses_tracked` says whether they mean anything here.
- **Recurring events.** `create_event` and `draft_meeting` take `recurrence`: `daily`,
  `weekly`, `weekdays`, `monthly`, `yearly`, or JSON with `interval`, `days`, `until` or
  `count`. `get_event` describes a series (`recurrence`, `recurrence_summary`). `update_event`
  changes one occurrence with `occurrence_date` and otherwise the whole series (the result
  says which). `delete_event` on a series now refuses unless told `occurrence_date` or
  `whole_series`; before, an occurrence's entry_id deleted the whole series without a word.
- **Find mail by attachment name.** `search_emails` takes `attachment_name`. Outlook cannot
  filter on it, so the search narrows to messages with attachments and reads each one's
  file names (up to 1,500 per folder, within the time limit); matches list their
  `attachments`. Pair it with a sender or a date range.

### Fixed
- Real attachments from Exchange Online senders were taken for inline pictures. Exchange
  stamps a Content-ID on every attachment, and the inline test treated any Content-ID as
  "embedded picture", so `draft_forward` with `include_thread_attachments` silently skipped a
  law firm's .docx and .pdf files and the new name search could not see them. The hidden
  property and the body-reference flag now decide; a Content-ID alone counts only for
  picture files.
- `create_event`, `draft_meeting` and `update_event` accept ISO 8601 with a "T" separator,
  as their descriptions always claimed; Outlook rejected it. Dates now reach Outlook as
  strings it parses itself: pywin32 converts a naive `datetime` from local time to UTC on
  the way into COM, and Outlook reads the result as wall time, so a 09:00 appointment
  landed at 16:00. (`GetOccurrence` needs a datetime, so it gets the wall time labelled
  UTC, which makes that conversion a no-op.)

## 0.10.0 - 2026-10-29

### Added
- **Inline answers.** `draft_reply` takes `inline_replies`, a list of `{"after": <phrase from
  the original>, "text": <answer>}`: each answer is written inside the quoted message, right
  under the paragraph or list item that contains the phrase, in the new `inline_color` setting
  (dark red by default), the way people answer a numbered list point by point in Outlook.
  Anchors are matched on the text a reader sees, so Word's tags, entities and curly quotes in
  between do not matter, and a phrase that repeats ("Confirm ok.") resolves in document order.
  An anchor that is not found saves nothing and is named in the result. `update_draft` takes
  the same parameter and rewrites the set. Until now the only way to answer inline was an HTML
  file pasted by hand.
- **Flagged mail.** Every email summary carries `flagged` (a follow-up flag not yet complete)
  and, when flagged, `flag_due`. `list_emails` and `search_emails` take `flagged_only`; it may
  stand alone in `search_emails` for "my flagged emails". Outlook stores a modern flag as a task
  on the message, which `[FlagStatus]` filters do not see, so the filter uses the task's
  complete property.

### Fixed
- Drafts the server saved showed in Outlook's reading pane without the editor (a bare header
  with "None" in it, no From or Send) until the user opened and closed them once. The hidden
  inspector `GetInspector` creates to fill the reply was never closed, and the draft was saved
  unread where Outlook's own are read. `_save_draft` now marks the draft read and closes the
  inspector after saving.

## 0.9.2 - 2026-10-27

### Changed
- `health_check` no longer reads like a warning about body-text search. `body_search_index`
  is now a plain status ("indexed" or "not indexed", with the probe word) rather than a
  sentence about slow scans, and the tool description says it and `legacy_mode` are
  diagnostics to mention only when the user asks about search speed. A version question got
  a lecture about search indexing; it should not.

## 0.9.1 - 2026-10-27

### Changed
- Wording only. `read_email` and `get_conversation` now say to call `reply_context` as well
  when the user wants to reply or forward: those are the tools Claude loads to read a thread,
  so they were the only descriptions in front of it while it drafted, and neither mentioned
  it. The server instructions, `reply_context`, `draft_reply` and `draft_forward` now say to
  skip the Context block when the thread has already been summarised in the conversation, so
  the block appears once per thread rather than above every draft; the To/CC from
  `reply_context` are still to be used every time.

## 0.9.0 - 2026-10-27

### Added
- **`draft_forward_event`**: forward a calendar event without sending. For a meeting the user
  organises, the recipients are added as attendees and the meeting is saved for the user to
  press Send Update (Outlook offers to send it only to them). For a received meeting or a
  plain appointment, an email with the event attached as .ics, saved to Drafts. Outlook's
  automation has no true "forward invitation" for appointments; the result says which form
  was used.
- **`move_event`**: move an event to another calendar in the same mailbox or in another
  account, one call per event. The result says when the destination will not pop reminders
  (secondary calendars) and when a meeting moved across accounts has become a plain copy that
  the organiser's updates will not reach.
- **Categories per account.** `list_categories` now reads the chosen account's own list (each
  mailbox keeps one; the old code read the profile-wide list whatever account was asked for)
  and names each colour. `create_event`, `draft_meeting` and `update_event` take
  `categories`; `set_category` and all three report `unknown_categories` when a name is not on
  that account's list (Outlook applies it uncoloured). Event and email summaries carry
  `categories`.

### Fixed
- `create_event` without an account went to Outlook's own default mailbox, ignoring the
  `default_account` setting. It now uses the default account like every other tool.

### Changed
- `create_event` and `draft_meeting` say plainly that the server cannot add a Teams link and
  how the user adds one (the Teams Meeting button in the window that opens, or Outlook's "Add
  online meeting to all meetings" option). README and CLAUDE.md say the same.

## 0.8.2 - 2026-10-27

### Fixed
- `contact_history` reported `sent_scan_capped` whenever an account's Sent Items walk examined
  its 100 messages without reaching the requested count, even when every relevant message had
  been found or the window held no more messages. It is now reported only when the walk
  stopped with messages still unread, and `sent_scan_partial` names each such account with how
  many of its sent messages were checked and why it stopped.

## 0.8.1 - 2026-10-27

### Changed
- `reply_context` now counts the messages that come after the one being answered
  (`newer_messages`, `newest_entry_id`) and, when there are any, says in its note to reply
  to the newest unless the user means this one and to tell the user. Before, the result only
  carried `requested_is_newest: false` and left it to the model to notice.

## 0.8.0 - 2026-10-27

Prompted by a Claude Desktop session that took two minutes to find one thread with a company:
nine calls in four rounds, four of them `contact_history` sweeps that found nothing. Replaying
those nine calls read-only: 60.5 s of server time before, 43.8 s after, and the first call now
returns the thread on its own. The one-call form the instructions now steer to,
`search_emails(query="Earl Gray", folder="all", account="all")`, takes 1.2 s.

### Fixed
- **`contact_history` missed mail where the person was only a recipient.** The Inbox tree was
  matched on the sender alone, so a colleague writing to the person with the user copied, and
  the user's own replies filed outside Sent Items, were invisible; a search for the company's
  domain returned nothing although the thread was addressed to it. The To/CC/BCC display
  strings are now matched too, and each hit is tagged `received`, `sent` or `copied` from its
  sender and the account's own addresses (`copied_count` is new).
- **`contact_history` spent 30 s confirming "nothing" in Sent Items.** The Python fallback that
  resolves every recipient when an address-like needle finds nothing natively is now skipped
  when the Inbox tree already showed the address raw, and otherwise capped at 100 messages and
  3 s per account, 6 s per call, reported as `sent_scan_capped`. Legacy mode keeps the old walk.

### Added
- **`search_emails` `match="any"`** (default): the query is looked for in the subject, the
  sender and the recipients in one native filter; `match="subject"` is the old behaviour.
  **`folder="all"`** on `search_emails` and `list_emails` covers the Inbox tree, Sent Items and
  Archive in one call. "My history with company X" is now one call instead of two or more.
- **`contact_history` falls back to subject matches.** When nothing matches a person, the
  result carries `subject_matches` from one cheap subject pass, so a company, deal or project
  name given to the wrong tool still finds its thread.
- **Timing.** `contact_history` and `get_conversation` always report `elapsed_s`; a
  `search_emails`, `list_emails` or `search_events` result that took more than two seconds
  comes back as an object with `elapsed_s`; and the server logs one line per call with the
  tool's name and duration, so a slow session can be explained from Claude's MCP log instead
  of guessed at.
- `tests/perf_check.py` replays the company-name case (`--company`, `--domain`).

### Changed
- Server instructions: a company, deal or project name is a subject search, never
  `contact_history`; start in the default account; use one spelling and widen rather than
  trying variants; at most one confirming call once the answer is in hand.
- Docs checked against the running server: README badges no longer point at upstream's PyPI
  package; the register-from-source command matches SETUP.md; CLAUDE.md and SETUP.md say the
  installer adds Git when missing; CONTRIBUTING.md describes the workflows that actually run.

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
