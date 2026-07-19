EXPENSE TRACKER
===============

HOW TO USE
1. Put your statements in the "imports" folder.
   - Amex: the CSV export (filename does not matter).
   - CIBC: the PDF statement (filename does not matter).
   - Add as many files as you like. Re-exported / overlapping statements are fine;
     duplicate transactions are removed automatically.
2. Run it:
   - Mac:      double-click  run.command
   - Windows:  double-click  run.bat
3. The first time it sees a new merchant it doesn't recognize, it asks you which
   category it belongs to (type a number, or type a brand-new category name, or
   press Enter to skip, or 'q' to stop for now). Your answers are saved and
   remembered, so it only ever asks once per merchant. Merchants are listed
   biggest-spend first, so you can categorize the important ones and press 'q'.
4. Open expense_tracker.xlsx and go to the Dashboard tab.
   The top line shows the date and time it was last updated.

FIRST RUN
The first run sets itself up automatically (needs internet, about a minute).
Every run after that is fast and works offline.
Requires Python 3 on the machine: https://www.python.org/downloads/
(On Windows, tick "Add python.exe to PATH" during install.)

MAC NOTE
The first time you open run.command, macOS may block it because it came from a zip.
Right-click run.command  ->  Open  ->  Open. You only do this once.

BEFORE RE-RUNNING
Close expense_tracker.xlsx in Excel first. If it is open, the tracker will save a
timestamped copy instead so nothing is lost.

YOUR CATEGORIES
Your merchant-to-category choices live in merchant_map.json. It is created and updated
automatically as you answer the prompts. Keep it with the tool. To re-do a merchant,
delete its line from that file (or delete the whole file to start categorizing fresh).