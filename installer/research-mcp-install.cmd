@echo off
rem research-mcp installer for Windows: double-click this file.
rem It installs uv (Astral's Python tool runner, no admin needed) if it is missing, then runs
rem "uvx research-mcp@latest setup", which checks Word and Zotero and adds research-mcp to the
rem ChatGPT app, Claude, Codex, Cursor and VS Code. Nothing else is downloaded.
setlocal
set "rc=1"
title research-mcp setup
echo research-mcp setup
echo.

where uvx >nul 2>nul
if not errorlevel 1 goto run

echo uv is not installed. Installing it from astral.sh for this user only...
powershell -NoProfile -ExecutionPolicy ByPass -Command "irm https://astral.sh/uv/install.ps1 | iex"
set "PATH=%USERPROFILE%\.local\bin;%USERPROFILE%\.cargo\bin;%PATH%"
where uvx >nul 2>nul
if not errorlevel 1 goto run
echo.
echo The uv install did not finish. Install it by hand from https://docs.astral.sh/uv/
echo then double-click this file again.
goto done

:run
echo.
uvx research-mcp@latest setup
set "rc=%errorlevel%"
if errorlevel 1 (
  echo.
  echo Setup ended with a problem: see the FAIL lines above. Run this file again after fixing them.
) else (
  echo.
  echo Done. Restart your apps: quit them from the system tray, then open them again.
)

:done
echo.
pause
exit /b %rc%
