# Manual setup

> **Most people should not need this page.** The README's *Getting started* has Claude do the
> whole install through Claude Code in the desktop app, and `install.cmd` (double-click) does the same
> without Claude. This page is the step-by-step equivalent for doing it by hand, for Claude
> Code's terminal, and for troubleshooting.

This gives Claude a live, local connection to your classic Outlook. Nothing leaves your PC except
what Claude reads into the conversation, exactly as if you had pasted it. Claude can never send
anything: at most it writes drafts, which open in Outlook for you to check and send yourself.

Allow about ten minutes. This assumes you use both Claude Code (the VS Code extension or the
`claude` command line) and the Claude desktop app, and connects the server to each. The install
in step 1 is shared; steps 2 and 3 just tell each app where it is.

## Before you start

| Need | Check |
|---|---|
| Windows 10/11 with **classic Outlook** installed and signed in | The "New Outlook" toggle (top right of Outlook) must be **off**. The new Outlook has no automation interface. |
| **Python 3.12 or newer** | Open PowerShell and run `python --version`. If it is missing or older, install from https://www.python.org/downloads/windows/ and tick **"Add python.exe to PATH"** in the installer. |
| **Claude Code** and the **Claude desktop app**, signed in | Claude Code: https://claude.com/claude-code. Desktop app: https://claude.ai/download or the Microsoft Store. Local MCP servers need the **Developer** section in the app's Settings, which free plans do not have. |
| **Git** (optional) | Only to clone. You can also download the repo as a zip from GitHub and unzip it. |

## 1. Get the code and install it (once)

Open PowerShell (Start menu, type `PowerShell`, press Enter) and run these one line at a time.
They put the code in a folder called `outlook-desktop-mcp` inside your user folder, so the path
is `C:\Users\<your username>\outlook-desktop-mcp`. You will need that path in steps 2 and 3.

```powershell
cd $HOME
git clone https://github.com/jfeuerman24/outlook-desktop-mcp.git
cd outlook-desktop-mcp
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -e .
dir .\.venv\Scripts\outlook-desktop-mcp.exe
```

The last line must list the exe. That file is the server; the rest of this page is about telling
Claude where it is. It runs the code in this folder directly, so if you ever change the code there
is nothing to reinstall.

To see your username, run `echo $env:USERNAME` in the same PowerShell window.

If PowerShell refuses to run scripts, you do not need `Activate.ps1`; the commands above call the
venv's Python directly. On an ARM64 PC, `pip install` may fail on a package called `cryptography`;
the README's ARM64 section has the two-line workaround.

## 2. Connect it to Claude Code (VS Code extension or terminal)

Claude Code in VS Code and the `claude` command share one list of MCP servers. Add this server
once, for every project, by running this in a terminal (the VS Code terminal panel is fine), with
the path from step 1:

```powershell
claude mcp add --scope user outlook-desktop -- C:\Users\<your username>\outlook-desktop-mcp\.venv\Scripts\outlook-desktop-mcp.exe
```

Then start a new Claude Code conversation. Type `/mcp` in the chat and **outlook-desktop** should
be listed as connected. Try "List my Outlook accounts."

To take it out again: `claude mcp remove outlook-desktop`.

## 3. Connect it to the Claude desktop app

The desktop app keeps a small settings file that lists the local servers it should start. You are
going to open that file in Notepad and add six lines to it.

**a. Find your exe path first.** Open File Explorer and go to the folder from step 1
(`C:\Users\<your username>\outlook-desktop-mcp`), then into `.venv`, then `Scripts`. You should see
`outlook-desktop-mcp.exe`. Click once in the address bar at the top of File Explorer, and it turns
into text like `C:\Users\yourname\outlook-desktop-mcp\.venv\Scripts`. Write that down or leave the
window open. Your full exe path is that text plus `\outlook-desktop-mcp.exe`.

**b. Open the settings file.**

1. Open the Claude desktop app.
2. Click your initials (or profile picture) in the **bottom-left** corner, then **Settings**.
3. In the left-hand list inside Settings, click **Developer**.
   (If there is no **Developer** entry, your plan does not include local MCP servers. Stop here.)
4. Click **Edit Config**. A File Explorer window opens with a file called
   `claude_desktop_config.json` selected. Nothing else happens; that is expected.
5. Right-click `claude_desktop_config.json` → **Open with** → **Notepad**.
   (If Notepad is not offered, choose **Choose another app** → **Notepad**.)

**c. Put the server in the file.** What you do depends on what Notepad shows.

*If the file is empty, or contains only `{}` or `{ "mcpServers": {} }`:* select everything
(Ctrl+A), delete it, and paste this in:

```json
{
  "mcpServers": {
    "outlook-desktop": {
      "command": "C:\\Users\\yourname\\outlook-desktop-mcp\\.venv\\Scripts\\outlook-desktop-mcp.exe"
    }
  }
}
```

*If the file already lists other servers* (you will see `"mcpServers": {` followed by one or more
named blocks), do not replace it. Find the closing `}` of the last server block, type a comma
after it, press Enter, and paste only this part:

```json
    "outlook-desktop": {
      "command": "C:\\Users\\yourname\\outlook-desktop-mcp\\.venv\\Scripts\\outlook-desktop-mcp.exe"
    }
```

**d. Fix the path.** Change `yourname` to your own username so the line matches the exe path you
found in (a). Two rules:

- Every single backslash in the path becomes a **double** backslash: `C:\Users\yourname\...`
  is written `C:\\Users\\yourname\\...`. The file will not load otherwise.
- Keep the quotation marks exactly as shown.

**e. Save and restart.**

1. Press Ctrl+S in Notepad, then close it.
2. **Fully quit Claude.** Closing the window is not enough: right-click the Claude icon in the
   system tray (bottom-right, next to the clock; click the small `^` arrow if it is hidden) and
   choose **Quit**.
3. Make sure classic Outlook is open.
4. Start Claude again.

**f. Check it worked.** Open **Settings → Developer** again. **outlook-desktop** should be listed
with a running status. In a new conversation, the button under the message box that opens
connectors and tools should also list **outlook-desktop**. If it shows an error instead, the path
in the file has a typo (compare it character by character with the address bar from step a) or
Outlook was not open when Claude started.

The first time Claude uses one of the tools in a conversation, the app asks you to allow it. That
prompt is Claude's, and is separate from Outlook's own prompt described next.

## 4. Outlook's security prompt

The first time Claude calls a tool, Outlook may show *"A program is trying to access email
address information stored in Outlook."* Either:

- tick **Allow access for** and pick **10 minutes** (you will see it again in later sessions), or
- suppress it permanently: right-click `docs\suppress-outlook-oom-prompts.reg` in the folder from
  step 1 → **Merge** (say yes to the admin prompt), then quit Outlook completely and reopen it.
  Details and caveats are in the README under *Programmatic Access security prompts*.

On ARM64 Windows, the Trust Center "antivirus status" route described in the README does not work;
the registry file is the only permanent fix there.

## 5. Try it

In a new Claude conversation:

- "List my Outlook accounts."
- "List my calendars."
- "What's on all my calendars Tuesday and Wednesday?"
- "Show me emails from Jordan in the last two days, including subfolders."
- "Read the newest one and give me the whole thread."
- "Draft a reply saying I'll have the numbers by Thursday." The reply opens in Outlook; read it
  and press Send. Claude cannot send it for you.

## If something is off

| Symptom | Fix |
|---|---|
| `python` is not recognised | Reinstall Python with "Add python.exe to PATH" ticked, or use the full path to `python.exe`. |
| `pip install` fails on ARM64 Windows | See the README's ARM64 notes: install `pywin32 mcp` with `--only-binary=:all:` first. |
| In Claude Code, `/mcp` shows outlook-desktop as failed | Outlook is not running, or the security prompt is waiting behind another window. To see the server's own startup message, run the exe by hand in a terminal, read what it prints, then press Ctrl+C. |
| No **Developer** section in the desktop app's Settings | Your plan does not include local MCP servers. |
| **outlook-desktop** never appears in the desktop app | Usually a path typo in the config (step 3d: every backslash doubled, username correct). **Settings** → **Developer** lists the server with its status and has an **Open Logs Folder** button; read the file whose name starts with `mcp-server-outlook-desktop`. |
| Tools appear but every call errors | Outlook is not running, "New Outlook" is on, or the security prompt is waiting behind another window. |
| You changed the code and Claude still behaves the old way | The server starts when Claude does. Start a new Claude Code session, or quit and reopen the desktop app. |

## Self-test

With Outlook open, this lists your mailboxes, calendars and a recent thread. It reads only and
changes nothing. Run it from the folder you installed into:

```powershell
.\.venv\Scripts\python.exe tests\live_check.py
```

## Draft settings

New drafts are saved to Drafts without opening a window. To have them pop open in Outlook,
add `"env": { "OUTLOOK_MCP_OPEN_DRAFTS": "1" }` inside the `"outlook-desktop"` block of the
desktop app's config (or `-e OUTLOOK_MCP_OPEN_DRAFTS=1` on the `claude mcp add` command).

### Font

Drafts written through this server are in **Aptos 12 pt black** by default, with your usual
signature underneath. To change the font, set environment variables before Claude starts the
server. In the desktop app's config that is an extra line inside the `"outlook-desktop"` block:

```json
"env": { "OUTLOOK_MCP_FONT_FAMILY": "Calibri", "OUTLOOK_MCP_FONT_SIZE_PT": "11", "OUTLOOK_MCP_FONT_COLOR": "#000000" }
```

For Claude Code, add `-e OUTLOOK_MCP_FONT_FAMILY=Calibri` (and so on) to the `claude mcp add` command.
