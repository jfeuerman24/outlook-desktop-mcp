# Guide for Claude: setting up and explaining this Outlook connector

This file is for Claude. It is read automatically when this folder is open in Claude Code
(inside the Claude desktop app). It says how to take a person from "I have this folder" to "Claude can
read my Outlook", and how to explain the tool along the way.

If a person is reading this: in the Claude desktop app, use Claude Code to open this folder on
this computer and say **"let's get started"**. Claude does the rest.

## What this is, in plain terms

- It lets Claude use the Outlook that is already open on this PC: read mail, calendars and
  tasks, pull a whole email thread together, find things, and write drafts.
- **Nothing is ever sent.** Email drafts wait in the Drafts folder and unsent meetings open in
  an Outlook window, for the person to read and press Send. There is no tool that sends,
  accepts or declines an invitation, or cancels a meeting.
- Nothing leaves the PC except what Claude reads into the conversation, exactly as if the
  person had pasted the email in.
- No passwords, no Microsoft account setup, no IT approval. It talks to the Outlook program the
  way an add-in does, so it has the access the person already has, only while Outlook is open.
- Under the hood it is a small program (an "MCP server") that Claude starts in the background.
  Claude asks it "list the accounts", it asks Outlook, and passes the answer back.

## When the person says "let's get started" (or anything like it)

Work through the stages in order. Keep messages short and plain; assume no technical
background. The person never opens PowerShell and never edits a file by hand. You run the
commands. Ask before each stage that changes something.

**First, check where you are running.** Run `uname` (or look at the shell). If it says
`Linux`, this folder is attached to an ordinary chat and your shell is a sandbox that cannot run
Windows programs. Tell the person, in one line, to open this same folder with **Claude Code** in the
Claude desktop app instead (working on this computer), then say "let's get started" there.
Alternatively they can double-click `install.cmd` in the folder, wait for "Done", and come back.
Do not try to work around this.

### Stage 1: Explain

Before touching anything, give the five bullets from "What this is" above in your own words,
then say setup takes about five minutes, you will do all of it, and they will need to quit and
reopen Claude once at the end. Ask if they are ready.

### Stage 2: Check the PC

Run:

    powershell.exe -NoProfile -ExecutionPolicy Bypass -File install.ps1 -CheckOnly

It prints `[OK]`, `[WARN]` or `[MISSING]` lines for Windows, classic Outlook, Python and the
Claude settings folder. Summarise in two or three sentences. If classic Outlook is `[MISSING]`,
stop: this tool cannot work with the new Outlook app or with Outlook on the web. If Python is
missing, say the installer will add it (Windows may ask for permission once). Ask the person to
confirm Outlook is open and that the **New Outlook** switch in its top-right corner is off.

Then ask which mailboxes they expect Claude to use. Every one has to be an account in Outlook
on this PC: **File → Add Account** for a mailbox of their own; shared mailboxes they have full
access to appear on their own, otherwise **File → Account Settings → Account Settings**, select
the account, **Change → More Settings → Advanced → Add**. If one is missing, have them add it
now, before the install, so the test in Stage 5 shows it. Note the names they give; Stage 5
compares them with what Outlook reports.

### Stage 3: Install

Run:

    powershell.exe -NoProfile -ExecutionPolicy Bypass -File install.ps1

It takes one to three minutes. It creates `.venv` in this folder, installs the server into it,
and then adds the server to Claude's own settings file (`claude_desktop_config.json`, wherever this
installation of Claude keeps it; the script prints the path), keeping everything already in that file
and saving a timestamped backup first. Tell the person
what it did in one or two sentences. If it fails, read the error, fix what is obvious (most often
Outlook not installed, or no internet for the downloads), and run it again; it is safe to repeat.
Do not improvise other install commands. `SETUP.md` has the manual steps if all else fails.

### Stage 4: Restart Claude

Claude only reads its settings file when it starts, so the person must restart it. Say exactly:

1. Right-click the Claude icon in the system tray (bottom-right, by the clock; click the small
   `^` arrow if it is hidden) and choose **Quit**. Closing the window is not enough.
2. Make sure classic Outlook is open.
3. Start Claude, come back to this folder, and say **"test it"**.

Warn them that this conversation will pause until they are back.

### Stage 5: Test, and Outlook's permission prompt

When they return, call the `list_accounts` tool (named something like
`mcp__outlook-desktop__list_accounts`). If the Outlook tools are not in your tool list, go to
Troubleshooting below.

The first call makes Outlook show *"A program is trying to access email address information
stored in Outlook"*. Tell the person this is Outlook's own safety check, not an error: tick
**Allow access for** and pick **10 minutes**. They will see it again in later sessions; that is
normal. There are two permanent fixes in README.md under *Programmatic Access security prompts*.
Mention they exist, but do not apply the registry file unless asked: it needs administrator
rights and it switches the check off for every program on the PC, not just this one.

Show them the accounts that came back and compare them with the mailboxes they named in
Stage 2. One that is missing is not in Outlook yet: give the Add Account steps from Stage 2 and
offer to check again once it is added. Say which account is the default: tools that are not
given an account read from it. If the person has several mailboxes, ask which one they want as
the default and set it with `set_setting("default_account", "<name>")`; it applies at once. Offer
to hide mailboxes they never use (`set_setting("hide_accounts", "...")`) so "all accounts"
searches skip them. Explain that settings can be changed any time by asking.

Accounts flagged `shared_uncached: true` are shared mailboxes Outlook keeps no local copy of.
Searching one means a server round trip per folder, so "all accounts" searches skip them on
their own (they can still be named). Say so in one sentence and, if the person never uses one,
offer to hide it as above. If they do need it in every search,
`set_setting("include_shared_in_all", "true")` turns that back on, at a cost of tens of seconds
per search.

### Stage 6: Show what it can do

Give five or six examples tuned to the accounts you just saw, for instance:

- "Show me unread email in my work inbox from the last two days."
- "Read the newest one and give me the whole thread."
- "What's on all my calendars Tuesday and Wednesday?"
- "What does the lease in that email say about assignment?" Claude reads the PDF with
  `read_attachment`; nothing is saved or opened.
- "When did we last hear from Chad?" `contact_history` covers every account and calendar.
- "Draft a reply saying I can do Thursday." You first show a few lines about the thread and the
  recipients Outlook would use, then the reply; once they agree it lands in Drafts, where they
  open it and press Send.
- "Open that email in Outlook." `open_in_outlook` shows it in Outlook's own window.
- "Set up a 30-minute meeting with alex@example.com tomorrow at 2pm." The invitation opens in an
  Outlook window, unsent, for them to review and send.

Repeat once that nothing is sent without them pressing Send in Outlook. Point to README.md for
the full list of tools.

## When the person says "update it" (or "get the latest version")

1. Check that a `.git` folder exists and run `git status --short`. If there are local edits, say
   so and ask before going on (`git stash` keeps them).
2. Run `git pull --ff-only`. If there is no `.git` folder the code was unzipped rather than
   cloned; say that the next step replaces the folder's files with the latest version (`.venv`
   and their settings are untouched), then run `git init`, `git remote add origin
   https://github.com/jfeuerman24/outlook-desktop-mcp.git`, `git fetch origin` and
   `git checkout -f -B main origin/main`.
3. Run `powershell.exe -NoProfile -ExecutionPolicy Bypass -File install.ps1`. It adds any new
   library and leaves a working install alone.
4. Walk through Stage 4 (restart Claude). When they are back, call `health_check`, which
   reports the new version, and say in one line what changed (CHANGELOG.md has it).

Settings live in `%LOCALAPPDATA%\outlook-desktop-mcp\settings.json`, outside the folder, so an
update never changes them.

## Troubleshooting

| Symptom | What to do |
|---|---|
| Anything fails and the cause is not obvious | Call `health_check` first. It says whether Outlook is reachable, which accounts loaded, the settings in force, and whether Outlook's permission prompt is blocking address access (common when the laptop lid is closed and nobody can click Allow). |
| After the restart, no Outlook tools in your tool list | Run `.venv\Scripts\python.exe scripts\configure_claude_desktop.py --show` and check the entry and that the exe exists. Ask the person to check **Settings → Developer**: the server should be listed with a status; **Open Logs Folder** has `mcp-server-outlook-desktop.log`. Most often Claude was not fully quit, or Outlook was not running when Claude started. |
| `install.ps1` says it cannot find Claude's settings folder, or Claude does not pick up the server after a restart | Ask the person to click their initials (bottom-left) → **Settings** → **Developer** → **Edit Config**. A File Explorer window opens on `claude_desktop_config.json`; ask them to click the address bar and paste you the path. Then run `.venv\Scripts\python.exe scripts\configure_claude_desktop.py --config "<that path>\claude_desktop_config.json"` and have them restart Claude. |
| Tools exist but every call errors | Outlook is not running, the New Outlook switch is on, or Outlook's permission prompt is waiting behind another window. |
| A draft opened in a window the person did not want, or a meeting did not open | `set_setting("open_in_outlook", ...)` with `none`, `meetings` (the default: unsent invitations only) or `all`. It applies at once. `open_in_outlook` the tool opens any single item on request. |
| Searches are slow, or a result says `"partial": true` | Call `health_check`: accounts marked `shared_uncached` are skipped by all-account searches (hide them if unused), and the `body_search_index` line says whether body-text search has to scan every message (if so, prefer sender and subject searches and add a date range). A partial result hit the time limit: narrow the request rather than repeat it, or raise `call_timeout_s`. If a search looks wrong rather than slow, `set_setting("legacy_mode", "true")` runs the pre-0.6 code at once; say what happened so it can be fixed. |
| `install.ps1` cannot download packages | No internet, or a corporate proxy. Retry; otherwise `SETUP.md` has the ARM64 and proxy notes. |
| Person wants it gone | `.venv\Scripts\python.exe scripts\configure_claude_desktop.py --remove`, then they restart Claude. Deleting the folder removes the rest. |

## Rules

- Never run `docs\suppress-outlook-oom-prompts.reg` or anything else needing administrator
  rights without asking first and explaining what it does.
- Edit Claude's settings file only through `scripts\configure_claude_desktop.py`, which keeps a
  backup. Never hand-edit it.
- Do not tell the person this tool can send email. It cannot.
- Drafts are written in the chat first. When asked for an email, reply or meeting, show the
  whole thing in the conversation, revise it there, and then ask explicitly whether to put it
  in Outlook. Call `draft_email`, `draft_reply`, `draft_forward` or `draft_meeting` only after a yes.
- Context before a reply. The person often has not opened the email in Outlook. Before showing
  a reply or forward draft, call `reply_context` on the message and open the preview with a
  Context block of two to four lines (who is on the thread, what has been said, what the newest
  message asks), then the To/CC it returned, which are Outlook's own, then the text.
- Open an item in Outlook (`open_in_outlook`) only when asked. By default email drafts stay in
  Drafts without a window and unsent meetings open on their own; the setting of the same name
  changes that.
- A revision to a draft that already exists in Outlook goes through `update_draft` (same
  explicit yes first), never a second `draft_*` call that leaves two drafts behind.
- Do not read their mail during setup beyond the test call, unless they ask.
- Keep the chat about setup; this folder's code is not the point for a new user.
