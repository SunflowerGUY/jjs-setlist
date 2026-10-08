#!/bin/bash
# Starts JJ's Setlist (browser edition) on a Mac and opens it in your web browser.
# Double-click this file. A Terminal window opens: keep it open while you use
# the app (you can minimise it); close it to stop. Messages appear here.
# Settings and saved setlists are kept as files in this folder.

cd "$(dirname "$0")" || exit 1

pause_and_exit() {
    echo
    read -n 1 -s -r -p "Press any key to close..."
    echo
    exit 1
}

# Find Python 3.8+. A new Mac may only have Apple's placeholder "python3",
# which offers to install the Command Line Tools the first time it's used.
PY=""
for candidate in python3 /usr/local/bin/python3 /opt/homebrew/bin/python3 \
                 /Library/Frameworks/Python.framework/Versions/Current/bin/python3; do
    if command -v "$candidate" >/dev/null 2>&1 &&
       "$candidate" -c 'import sys; sys.exit(sys.version_info < (3, 8))' >/dev/null 2>&1; then
        PY="$candidate"
        break
    fi
done
if [ -z "$PY" ]; then
    echo "JJ's Setlist needs Python 3.8 or newer, and it wasn't found."
    echo
    echo "Either install it from https://www.python.org/downloads/macos/"
    echo "or, if macOS offered to install the \"command line developer tools\","
    echo "click Install, wait for it to finish, then double-click this file again."
    pause_and_exit
fi

"$PY" serve.py "$@" || pause_and_exit
