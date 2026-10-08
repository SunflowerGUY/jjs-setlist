@echo off
rem Makes "packages\JJs Setlist - Browser Edition.zip" - one zip for Windows, Mac and Linux.
rem   Make Package.bat            for your band (with the song spreadsheet, setlists and Drive link)
rem   Make Package.bat --public   the app only, no personal files
title Make Package
cd /d "%~dp0"
echo Making the JJ's Setlist package for Windows, Mac and Linux...
echo.
set "PY=python"
python -c "" >nul 2>nul || set "PY=py -3"
%PY% make_package.py %*
if errorlevel 1 (
    echo.
    echo *** FAILED - see the messages above. ***
) else (
    echo.
    echo Done.
)
echo.
pause
