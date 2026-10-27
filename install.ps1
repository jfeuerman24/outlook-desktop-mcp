<#
.SYNOPSIS
  Sets up outlook-desktop-mcp on this PC and registers it with the Claude desktop app.

.DESCRIPTION
  Does, in order, skipping anything already done:
    1. Checks this is Windows and that classic Outlook is installed.
    2. Finds Python 3.12 or newer. If there is none, installs it with winget.
    3. Creates .venv next to this script and installs the server into it.
    4. Adds the server to Claude's settings file (%APPDATA%\Claude\claude_desktop_config.json),
       keeping everything already in that file and saving a backup first.
  Nothing here sends mail, touches the mailbox, or needs administrator rights
  (winget may show one UAC prompt if Python has to be installed).

  Safe to run again at any time.

.PARAMETER CheckOnly
  Only report what is installed and what is missing; change nothing.

.PARAMETER SkipClaudeConfig
  Do steps 1-3 but leave Claude's settings file alone.

.EXAMPLE
  powershell -NoProfile -ExecutionPolicy Bypass -File install.ps1
  powershell -NoProfile -ExecutionPolicy Bypass -File install.ps1 -CheckOnly
#>
[CmdletBinding()]
param(
    [switch]$CheckOnly,
    [switch]$SkipClaudeConfig
)

$ErrorActionPreference = 'Stop'
$Root = $PSScriptRoot
$script:Problems = 0

function Ok($m)   { Write-Host "[OK]      $m" -ForegroundColor Green }
function Warn($m) { Write-Host "[WARN]    $m" -ForegroundColor Yellow }
function Miss($m) { Write-Host "[MISSING] $m" -ForegroundColor Red; $script:Problems++ }
function Step($m) { Write-Host ""; Write-Host "== $m" -ForegroundColor Cyan }

# Run a native command (pip, python, winget) and stop with a readable message if it fails.
# Windows PowerShell 5.1 turns anything a native command prints to stderr into a terminating
# error when $ErrorActionPreference is 'Stop', so native commands run with 'Continue' and are
# judged by their exit code instead.
function Run([string]$What, [scriptblock]$Cmd) {
    $prev = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try { & $Cmd } finally { $ErrorActionPreference = $prev }
    if ($LASTEXITCODE -ne 0) { throw "$What failed (exit code $LASTEXITCODE). See the output above." }
}
# Same, but failure is allowed and nothing is printed.
function Quiet([scriptblock]$Cmd) {
    $prev = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try { & $Cmd *> $null } catch { } finally { $ErrorActionPreference = $prev }
}

# ---------------------------------------------------------------- 1. Windows + Outlook
Step "Checking this PC"

if ($env:OS -ne 'Windows_NT') {
    Miss "This only runs on Windows. It drives the classic Outlook desktop app through COM, which exists nowhere else."
    exit 1
}
Ok "Windows, $env:PROCESSOR_ARCHITECTURE"

$outlookExe = $null
foreach ($key in @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE')) {
    if (Test-Path $key) {
        $v = (Get-ItemProperty $key).'(default)'
        if ($v -and (Test-Path $v)) { $outlookExe = $v; break }
    }
}
if ($outlookExe) {
    Ok "Classic Outlook: $outlookExe"
} else {
    Miss "Classic Outlook (OUTLOOK.EXE) was not found. The new Outlook app (olk.exe) has no automation interface, so this tool cannot work with it. Install classic Outlook from Microsoft 365 and run this again."
}

if (Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue) {
    Ok "Outlook is running"
} else {
    Warn "Outlook is not running. It must be open (and the 'New Outlook' switch off) whenever Claude uses the tools."
}

# Where the Claude desktop app keeps claude_desktop_config.json depends on how it was installed:
# Microsoft Store build: %LOCALAPPDATA%\Claude-Data\<package id>; claude.ai download: %APPDATA%\Claude;
# older Store builds: inside the package's LocalCache. scripts\configure_claude_desktop.py uses
# the same list, in the same order, and edits the first one that already holds the file.
$claudeDirs = @()
$claudeDirs += @(Get-ChildItem "$env:LOCALAPPDATA\Claude-Data" -Directory -ErrorAction SilentlyContinue | ForEach-Object FullName)
$d = Join-Path $env:APPDATA 'Claude'
if (Test-Path $d) { $claudeDirs += $d }
$claudeDirs += @(Get-ChildItem "$env:LOCALAPPDATA\Packages\*\LocalCache\Roaming\Claude" -Directory -ErrorAction SilentlyContinue | ForEach-Object FullName)
$withFile = @($claudeDirs | Where-Object { Test-Path (Join-Path $_ 'claude_desktop_config.json') })
if ($withFile.Count -gt 0) {
    Ok "Claude desktop app settings file: $(Join-Path $withFile[0] 'claude_desktop_config.json')"
} elseif ($claudeDirs.Count -gt 0) {
    Ok "Claude desktop app settings folder: $($claudeDirs[0]) (no settings file yet; one will be created)"
} else {
    Warn "Claude desktop app settings folder not found (looked in %LOCALAPPDATA%\Claude-Data, %APPDATA%\Claude and the Store package folders). Install the Claude desktop app and open it once before the last step."
}

# ---------------------------------------------------------------- 2. Python
Step "Looking for Python 3.12 or newer"

function Test-PythonCandidate([string]$Exe, [string[]]$PreArgs) {
    $prev = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $out = & $Exe @PreArgs -c "import sys; print(sys.version_info[0], sys.version_info[1])" 2>$null
        if ($LASTEXITCODE -ne 0 -or -not $out) { return $null }
        $parts = "$out".Trim().Split(' ')
        $maj = [int]$parts[0]; $min = [int]$parts[1]
        if ($maj -gt 3 -or ($maj -eq 3 -and $min -ge 12)) {
            return [pscustomobject]@{ Exe = $Exe; PreArgs = $PreArgs; Version = "$maj.$min" }
        }
    } catch { } finally { $ErrorActionPreference = $prev }
    return $null
}

function Find-Python {
    $cands = New-Object System.Collections.ArrayList
    # Direct installs first: they work even when PATH has not been refreshed yet.
    foreach ($pat in @(
        "$env:LOCALAPPDATA\Programs\Python\Python31*\python.exe",
        "$env:LOCALAPPDATA\Programs\Python\Python31*-arm64\python.exe",
        "$env:ProgramFiles\Python31*\python.exe",
        "$env:ProgramFiles\Python31*-arm64\python.exe")) {
        Get-ChildItem $pat -ErrorAction SilentlyContinue | Sort-Object FullName -Descending |
            ForEach-Object { [void]$cands.Add(@{ Exe = $_.FullName; PreArgs = @() }) }
    }
    if (Get-Command py -ErrorAction SilentlyContinue) {
        foreach ($v in '3.14', '3.13', '3.12') { [void]$cands.Add(@{ Exe = 'py'; PreArgs = @("-$v") }) }
    }
    foreach ($name in 'python', 'python3') {
        if (Get-Command $name -ErrorAction SilentlyContinue) { [void]$cands.Add(@{ Exe = $name; PreArgs = @() }) }
    }
    foreach ($c in $cands) {
        $hit = Test-PythonCandidate $c.Exe $c.PreArgs
        if ($hit) { return $hit }
    }
    return $null
}

$py = Find-Python
if ($py) {
    Ok "Python $($py.Version): $($py.Exe) $($py.PreArgs -join ' ')"
} elseif ($CheckOnly) {
    Miss "No Python 3.12+ found. install.ps1 will install it with winget."
} else {
    Warn "No Python 3.12+ found. Installing it with winget (Windows may ask for permission once)..."
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Miss "winget is not available. Install Python 3.12 or newer from https://www.python.org/downloads/windows/ (tick 'Add python.exe to PATH'), then run this again."
        exit 1
    }
    $prev = $ErrorActionPreference; $ErrorActionPreference = 'Continue'
    try { & winget install -e --id Python.Python.3.12 --accept-source-agreements --accept-package-agreements --silent }
    finally { $ErrorActionPreference = $prev }
    # winget returns non-zero for "already installed" too, so just look again.
    $py = Find-Python
    if ($py) {
        Ok "Python $($py.Version) installed: $($py.Exe)"
    } else {
        Miss "Python still not found after winget. Install it from https://www.python.org/downloads/windows/ (tick 'Add python.exe to PATH'), then run this again."
        exit 1
    }
}

if ($CheckOnly) {
    Step "Check complete"
    $venvExe = Join-Path $Root '.venv\Scripts\outlook-desktop-mcp.exe'
    if (Test-Path $venvExe) { Ok "Server already installed: $venvExe" } else { Warn "Server not installed yet (no .venv). install.ps1 will create it." }
    if ($script:Problems -gt 0) { Write-Host "$($script:Problems) problem(s) above need attention."; exit 1 }
    Write-Host "Everything needed is present."
    exit 0
}

if ($script:Problems -gt 0) {
    Write-Host ""
    Write-Host "Stopping: fix the [MISSING] item(s) above and run this again." -ForegroundColor Red
    exit 1
}

# ---------------------------------------------------------------- 3. venv + packages
Step "Installing the server into $Root\.venv"

$venvDir = Join-Path $Root '.venv'
$venvPy  = Join-Path $venvDir 'Scripts\python.exe'
if (-not (Test-Path $venvPy)) {
    $pyExe = $py.Exe; $pyPre = @($py.PreArgs)
    Run "Creating the virtual environment" { & $pyExe @pyPre -m venv $venvDir }
    Ok "Created .venv"
} else {
    Ok ".venv already exists"
}

$exe = Join-Path $venvDir 'Scripts\outlook-desktop-mcp.exe'

# Already installed and working? Then leave it alone. Reinstalling over a server that Claude
# is currently running fails (Windows will not overwrite a program in use), and there is
# nothing to gain. Delete .venv to force a clean reinstall.
$alreadyInstalled = $false
if (Test-Path $exe) {
    $prev = $ErrorActionPreference; $ErrorActionPreference = 'Continue'
    try {
        $importOut = & $venvPy -c "import win32com.client, mcp, outlook_desktop_mcp; print('ok')" 2>&1
        $alreadyInstalled = ($LASTEXITCODE -eq 0)
        if (-not $alreadyInstalled) {
            Warn "Existing install in .venv does not load cleanly, so it will be reinstalled. Python said:"
            ($importOut | Out-String).Trim() -split "`n" | ForEach-Object { Write-Host "          $_" -ForegroundColor DarkGray }
        }
    } catch { } finally { $ErrorActionPreference = $prev }
}

if ($alreadyInstalled) {
    Ok "Server already installed: $exe (delete the .venv folder to reinstall from scratch)"
} else {
    if (Get-Process -Name outlook-desktop-mcp -ErrorAction SilentlyContinue) {
        throw "outlook-desktop-mcp.exe is running (Claude is using it) so it cannot be replaced. Quit the Claude desktop app from the system tray, then run this again."
    }
    # Leftovers from interrupted pip upgrades only produce warnings; clear them.
    Get-ChildItem (Join-Path $venvDir 'Lib\site-packages') -Directory -Filter '~*' -ErrorAction SilentlyContinue |
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

    Run "Upgrading pip" { & $venvPy -m pip install --quiet --disable-pip-version-check --upgrade pip }
    # Wheels only: avoids compiling anything, which matters on ARM64 where some
    # packages have no prebuilt wheel for the versions pip would otherwise pick.
    Run "Installing pywin32 and mcp" { & $venvPy -m pip install --quiet --disable-pip-version-check --only-binary=:all: "pywin32>=310" "mcp>=1.26.0,<2" }
    Run "Installing outlook-desktop-mcp" { & $venvPy -m pip install --quiet --disable-pip-version-check --no-deps -e $Root }
    # Registers the pywin32 DLLs; harmless if it has already been done or is not needed.
    Quiet { & $venvPy -m pywin32_postinstall -install }
    Ok "Packages installed"

    Run "Verifying imports" { & $venvPy -c "import win32com.client, mcp, outlook_desktop_mcp; print('imports ok')" }
    if (-not (Test-Path $exe)) { throw "Expected $exe after install, but it is not there." }
    Ok "Server executable: $exe"
}

# ---------------------------------------------------------------- 4. Claude desktop config
if ($SkipClaudeConfig) {
    Step "Skipping Claude desktop settings (-SkipClaudeConfig)"
} else {
    Step "Registering the server with the Claude desktop app"
    $cfgScript = Join-Path $Root 'scripts\configure_claude_desktop.py'
    Run "Updating Claude's settings file (the server itself is installed)" { & $venvPy $cfgScript --exe $exe }
}

# ---------------------------------------------------------------- done
Step "Done"
Write-Host "Next:"
Write-Host "  1. Fully quit the Claude desktop app: right-click its icon in the system tray (by the clock) and choose Quit."
Write-Host "  2. Make sure classic Outlook is open."
Write-Host "  3. Start Claude again and ask it: List my Outlook accounts."
Write-Host "     Outlook may ask whether to allow access; choose 'Allow access for 10 minutes'."
