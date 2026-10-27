@echo off
REM Double-click this file to install outlook-desktop-mcp and register it with the
REM Claude desktop app. It just runs install.ps1 in this folder and keeps the window
REM open so you can read the result. Safe to run again.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1" %*
echo.
pause
