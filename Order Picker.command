#!/bin/bash
# Double-clickable launcher for the interactive order picker.
#
# Double-click this in Finder: it opens Terminal, then runs pick.py from this repo
# using the repo's virtualenv (.venv) if one exists, otherwise the system python3.
# The owner never has to type anything.

# A .command opened from Finder starts in the user's home dir, so jump to this
# script's own folder (the repo) first.
cd "$(dirname "$0")" || { echo "Could not find the app folder."; read -r; exit 1; }

if [ -x ".venv/bin/python" ]; then
    PY=".venv/bin/python"
else
    PY="python3"
fi

# Close the Terminal window this script is running in. We do NOT rely on the
# Terminal profile's "when the shell exits" setting (the owner's is set to keep
# the window open, which is what left a dead window behind — and Terminal then
# restored that window on the next launch, giving two windows). Instead we close
# our own window by matching its tty, so other Terminal windows are untouched.
# Backgrounding osascript and exiting first means the shell is already gone when
# the window closes, so Terminal never shows its "a process is running" prompt.
close_own_window() {
    local target_tty
    target_tty="$(tty)"
    if [ "$TERM_PROGRAM" = "Apple_Terminal" ]; then
        osascript <<OSA &
tell application "Terminal"
    repeat with w in windows
        repeat with t in tabs of w
            if tty of t is "$target_tty" then
                close w
                return
            end if
        end repeat
    end repeat
end tell
OSA
    fi
    exit 0
}

"$PY" pick.py
status=$?

# On an error, pause so the message can be read before the window closes.
if [ "$status" -ne 0 ]; then
    echo
    echo "The picker exited with an error (code $status). Check the logs/ folder."
    echo "Press Return to close this window."
    read -r
fi

close_own_window
