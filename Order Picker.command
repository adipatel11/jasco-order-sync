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

"$PY" pick.py
status=$?

# On a clean exit the window just closes; on an error, pause so it can be read.
if [ "$status" -ne 0 ]; then
    echo
    echo "The picker exited with an error (code $status). Check the logs/ folder."
    echo "Press Return to close this window."
    read -r
fi
