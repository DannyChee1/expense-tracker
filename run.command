#!/bin/bash
cd "$(dirname "$0")" || exit 1

echo "Expense Tracker"
echo "==============="

if ! command -v python3 >/dev/null 2>&1; then
    echo ""
    echo "Python 3 is required but was not found."
    echo "Install it from https://www.python.org/downloads/ then double-click this again."
    echo ""
    read -n1 -r -p "Press any key to close..."
    exit 1
fi

NEED_SETUP=0
[ -x ".venv/bin/python" ] || NEED_SETUP=1
if [ "$NEED_SETUP" = "0" ]; then
    .venv/bin/python -c "import pandas, openpyxl, pdfplumber" 2>/dev/null || NEED_SETUP=1
fi

if [ "$NEED_SETUP" = "1" ]; then
    echo "First-time setup: installing (needs internet, ~1 minute)..."
    rm -rf .venv
    python3 -m venv .venv || { echo "Could not create environment."; read -n1 -r -p "Press any key..."; exit 1; }
    .venv/bin/python -m pip install --quiet --upgrade pip
    .venv/bin/python -m pip install --quiet -r requirements.txt || { echo "Install failed."; read -n1 -r -p "Press any key..."; exit 1; }
fi

.venv/bin/python build_tracker.py

echo ""
read -n1 -r -p "Press any key to close..."
