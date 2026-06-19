@echo off
REM Double-clickable launcher for the interactive order picker (Windows).
REM
REM Double-click this in File Explorer: it runs pick.py from this repo using the
REM repo's virtualenv (.venv) if one exists, otherwise whatever `python` is on PATH.
REM The owner never has to type anything. This is the Windows twin of
REM "Order Picker.command" (macOS); both launch the same pick.py.

REM A .bat opened from Explorer can start in the user's home dir, so jump to this
REM script's own folder (the repo) first. %~dp0 is that folder, with a trailing slash.
cd /d "%~dp0"

if exist ".venv\Scripts\python.exe" (
    set "PY=.venv\Scripts\python.exe"
) else (
    set "PY=python"
)

"%PY%" pick.py
set "STATUS=%ERRORLEVEL%"

REM On a clean exit the window just closes; on an error, pause so it can be read.
if not "%STATUS%"=="0" (
    echo.
    echo The picker exited with an error ^(code %STATUS%^). Check the logs\ folder.
    pause
)
