# Setup: use this with the Claude desktop app (or Claude Code)

This gives Claude a live, local connection to your classic Outlook. Nothing leaves your PC
except what Claude reads into the conversation, exactly as if you had pasted it. Allow about
ten minutes.

## Before you start

| Need | Check |
|---|---|
| Windows 10/11 with **classic Outlook** installed and signed in | The "New Outlook" toggle (top right of Outlook) must be **off**. The new Outlook has no automation interface. |
| **Python 3.12 or newer** | Open PowerShell and run `python --version`. If it's missing or older, install from https://www.python.org/downloads/windows/ and tick **"Add python.exe to PATH"** in the installer. |
| **Claude desktop app** on a paid plan (Pro, Max or Team) | Local MCP servers need the **Developer** section in Settings; free accounts don't have it. Download: https://claude.ai/download |
| **Git** (optional) | Only to clone. You can also download the repo as a zip from GitHub and unzip it. |

## 1. Get the code and install it

Open PowerShell (Start menu, type `PowerShell`) and run, one line at a time. Change
`C:\Tools` to wherever you want it to live; avoid OneDrive folders.

```powershell
cd C:\Tools
git clone https://github.com/jfeuerman24/outlook-desktop-mcp.git
cd outlook-desktop-mcp
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -e .
dir .\.venv\Scripts\outlook-desktop-mcp.exe
```

The last line must list the exe. If PowerShell refuses to run scripts, you don't need
`Activate.ps1`; the commands above call the venv's Python directly.

Quick self-test with Outlook open (prints your mailboxes and calendars, sends nothing):

```powershell
.\.venv\Scripts\python.exe tests\live_readonly_check.py
```

## 2. Connect it to the Claude desktop app

1. Open Claude. Click your initials (bottom left) → **Settings** → **Developer** → **Edit Config**.
2. File Explorer opens on `claude_desktop_config.json` (it lives in `%APPDATA%\Claude`). Open it in Notepad.
3. Paste this, fixing the path to match step 1. Keep the double backslashes; JSON needs them.

   ```json
   {
     "mcpServers": {
       "outlook-desktop": {
         "command": "C:\\Tools\\outlook-desktop-mcp\\.venv\\Scripts\\outlook-desktop-mcp.exe",
         "args": ["--read-only"]
       }
     }
   }
   ```

   If the file already lists other servers, add only the `"outlook-desktop": { … }` block inside
   the existing `"mcpServers"` section, with a comma after the previous entry.

4. **Decide on read-only.** With `"--read-only"` Claude can list, search and read mail,
   calendars and tasks, and save attachments, but it cannot send, reply, move, delete or create
   anything. Remove that line to let Claude act on your mailbox too. Start read-only; you can
   change it later.
5. Save the file, then **fully quit Claude**: right-click its icon in the system tray (bottom
   right, next to the clock) → **Quit**. Closing the window is not enough.
6. Make sure Outlook is open, then start Claude again. Under the message box, the tools button
   (slider icon) should list **outlook-desktop**.

## 3. Outlook's security prompt

The first time Claude calls a tool, Outlook may show *"A program is trying to access email
address information stored in Outlook."* Either:

- tick **Allow access for** and pick **10 minutes** (you'll see it again in later sessions), or
- suppress it permanently: run the file `docs\suppress-outlook-oom-prompts.reg` from this
  folder as administrator (right-click → Merge, or `reg import` from an elevated prompt), quit
  Outlook completely, and reopen it. Details and caveats are in the README under
  *Programmatic Access security prompts*.

## 4. Try it

In a new Claude chat:

- "List my Outlook accounts."
- "List my calendars in the personal account."
- "What's on all my calendars Tuesday and Wednesday?"
- "Show me emails from Freestone in the last two days, including subfolders."
- "Read the newest one and give me the whole thread."

## If something is off

| Symptom | Fix |
|---|---|
| No **Developer** section in Claude's Settings | Your plan doesn't include local MCP servers. |
| `outlook-desktop` never appears | Path typo in the config. Claude's logs are in `%APPDATA%\Claude\logs`; open the file starting with `mcp-server-outlook-desktop`. |
| Tools appear but every call errors | Outlook isn't running, "New Outlook" is on, or the security prompt is waiting behind another window. |
| `python` is not recognised | Reinstall Python with "Add python.exe to PATH" ticked, or use the full path to `python.exe`. |
| Pip install fails on ARM64 Windows | See the README's ARM64 notes: install `pywin32 mcp` with `--only-binary=:all:` first. |

## Using it from Claude Code instead

```powershell
claude mcp add outlook-desktop -- C:\Tools\outlook-desktop-mcp\.venv\Scripts\outlook-desktop-mcp.exe --read-only
```

## Outgoing font

Emails sent or replied to through this server are written in **Aptos 12 pt black** by default.
To change it, set environment variables before Claude starts the server, for example in the
config above:

```json
"env": { "OUTLOOK_MCP_FONT_FAMILY": "Calibri", "OUTLOOK_MCP_FONT_SIZE_PT": "11", "OUTLOOK_MCP_FONT_COLOR": "#000000" }
```

(Only relevant when read-only is off.)
