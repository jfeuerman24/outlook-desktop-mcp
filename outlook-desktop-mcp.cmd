@echo off
REM Outlook Desktop MCP - Launcher
REM Usage:
REM   outlook-desktop-mcp.cmd mcp [--read-only]   Start the MCP server (stdio)
REM   outlook-desktop-mcp.cmd test                Run COM validation tests
REM   outlook-desktop-mcp.cmd unittest            Run the offline unit tests

setlocal

set PYTHON=%~dp0.venv\Scripts\python.exe

if not exist "%PYTHON%" (
    echo ERROR: Virtual environment not found. Run setup first. 1>&2
    exit /b 1
)

if "%1"=="mcp" (
    "%PYTHON%" -m outlook_desktop_mcp %2 %3 %4
) else if "%1"=="test" (
    "%PYTHON%" tests\phase1_com_test.py
) else if "%1"=="unittest" (
    cd /d "%~dp0" && "%PYTHON%" -m unittest discover -s tests -v
) else (
    echo Usage: outlook-desktop-mcp.cmd [mcp [--read-only]^|test^|unittest] 1>&2
    exit /b 1
)
