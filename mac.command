#!/bin/bash
cd "$(dirname "$0")/app" || exit 1

# Pause so output stays visible, then actually close the Terminal window.
# (macOS keeps the window open when a script just exits, so we close it ourselves.)
finish() {
    echo ""
    read -n 1 -s -r -p "Press any key to close this window..."
    ( osascript -e 'tell application "Terminal" to close front window' >/dev/null 2>&1 & )
    exit "${1:-0}"
}

echo "Expense Tracker"
echo "==============="

if ! command -v python3 >/dev/null 2>&1; then
    echo ""
    echo "Python 3 is required but was not found."
    echo "Install it from https://www.python.org/downloads/ then double-click this again."
    finish 1
fi

NEED_SETUP=0
[ -x ".venv/bin/python" ] || NEED_SETUP=1
if [ "$NEED_SETUP" = "0" ]; then
    .venv/bin/python -c "import pandas, openpyxl, pdfplumber" 2>/dev/null || NEED_SETUP=1
fi

if [ "$NEED_SETUP" = "1" ]; then
    echo "First-time setup: installing (needs internet, ~1 minute)..."
    rm -rf .venv
    python3 -m venv .venv || { echo "Could not create environment."; finish 1; }
    .venv/bin/python -m pip install --quiet --upgrade pip
    .venv/bin/python -m pip install --quiet -r requirements.txt || { echo "Install failed."; finish 1; }
fi

.venv/bin/python build_tracker.py

finish 0
