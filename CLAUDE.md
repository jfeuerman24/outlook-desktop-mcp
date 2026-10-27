# Guide for Claude: setting up and explaining this Outlook connector

This file is for Claude. It is read automatically when this folder is open in Claude Code
(inside the Claude desktop app). It says how to take a person from "I have this folder" to "Claude can
read my Outlook", and how to explain the tool along the way.

If a person is reading this: in the Claude desktop app, use Claude Code to open this folder on
this computer and say **"let's get started"**. Claude does the rest.

## What this is, in plain terms

- It lets Claude use the Outlook that is already open on this PC: read mail, calendars and
  tasks, pull a whole email thread together, find things, and write drafts.
- **Nothing is ever sent.** Drafts open in Outlook for the person to read and press Send. There
  is no tool that sends, accepts or declines an invitation, or cancels a meeting.
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

Show them the accounts that came back and say which one Outlook treats as the default. Tools
that are not given an account read from that default, so for anyone with several mailboxes,
name the account in requests or ask Claude to.

### Stage 6: Show what it can do

Give five or six examples tuned to the accounts you just saw, for instance:

- "Show me unread email in my work inbox from the last two days."
- "Read the newest one and give me the whole thread."
- "What's on all my calendars Tuesday and Wednesday?"
- "Save the attachment from that email to my Downloads folder."
- "Draft a reply saying I can do Thursday." The reply opens in Outlook; they read it and press Send.
- "Set up a 30-minute meeting with alex@example.com tomorrow at 2pm." The invitation opens
  unsent for them to review.

Repeat once that nothing is sent without them pressing Send in Outlook. Point to README.md for
the full list of tools.

## Troubleshooting

| Symptom | What to do |
|---|---|
| After the restart, no Outlook tools in your tool list | Run `.venv\Scripts\python.exe scripts\configure_claude_desktop.py --show` and check the entry and that the exe exists. Ask the person to check **Settings → Developer**: the server should be listed with a status; **Open Logs Folder** has `mcp-server-outlook-desktop.log`. Most often Claude was not fully quit, or Outlook was not running when Claude started. |
| `install.ps1` says it cannot find Claude's settings folder, or Claude does not pick up the server after a restart | Ask the person to click their initials (bottom-left) → **Settings** → **Developer** → **Edit Config**. A File Explorer window opens on `claude_desktop_config.json`; ask them to click the address bar and paste you the path. Then run `.venv\Scripts\python.exe scripts\configure_claude_desktop.py --config "<that path>\claude_desktop_config.json"` and have them restart Claude. |
| Tools exist but every call errors | Outlook is not running, the New Outlook switch is on, or Outlook's permission prompt is waiting behind another window. |
| `install.ps1` cannot download packages | No internet, or a corporate proxy. Retry; otherwise `SETUP.md` has the ARM64 and proxy notes. |
| Person wants it gone | `.venv\Scripts\python.exe scripts\configure_claude_desktop.py --remove`, then they restart Claude. Deleting the folder removes the rest. |

## Rules

- Never run `docs\suppress-outlook-oom-prompts.reg` or anything else needing administrator
  rights without asking first and explaining what it does.
- Edit Claude's settings file only through `scripts\configure_claude_desktop.py`, which keeps a
  backup. Never hand-edit it.
- Do not tell the person this tool can send email. It cannot.
- Do not read their mail during setup beyond the test call, unless they ask.
- Keep the chat about setup; this folder's code is not the point for a new user.
