@echo off
rem Starts JJ's Setlist (browser edition) and opens it in your web browser.
rem Settings and saved setlists are kept as files in this folder.
rem Keep this window open while you use the app; close it to stop.
title JJ's Setlist
cd /d "%~dp0"

rem Find Python 3: "python", else the Windows "py" launcher.
set "PY="
python -c "import sys; sys.exit(sys.version_info < (3, 8))" >nul 2>nul && set "PY=python"
if not defined PY (
    py -3 -c "import sys; sys.exit(sys.version_info < (3, 8))" >nul 2>nul && set "PY=py -3"
)
if not defined PY (
    echo JJ's Setlist needs Python 3.8 or newer, and it wasn't found.
    echo.
    echo Install it from https://www.python.org/downloads/
    echo ^(tick "Add python.exe to PATH" in the installer^), then double-click this file again.
    echo.
    pause
    exit /b 1
)

%PY% serve.py %*
if errorlevel 1 (
    echo.
    pause
)
