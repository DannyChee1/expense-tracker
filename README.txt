EXPENSE TRACKER
===============

FOLDER LAYOUT
  mac.command      <- Mac: double-click this
  windows.bat      <- Windows: double-click this
  imports/         <- put your statements here
  app/             <- the code + settings (you don't need to touch this)
  expense_tracker.xlsx   <- created here after you run it

HOW TO USE
1. Put your statements in the "imports" folder.
   - Amex: the CSV export.
   - CIBC: the PDF statement.
   - Filenames don't matter. Add as many files as you like. Re-exported or
     overlapping statements are fine - duplicate transactions are removed
     automatically. Any file it doesn't recognize is skipped (it tells you which).
2. Run it:
   - Mac:      double-click  mac.command
   - Windows:  double-click  windows.bat
3. The first time it sees a merchant it doesn't recognize, it asks which category
   it belongs to (type a number, type a brand-new category name, press Enter to
   skip, or 'q' to stop). Answers are saved, so it only asks once per merchant.
   Merchants are listed biggest-spend first.
4. Open expense_tracker.xlsx and go to the Dashboard tab. The top-left filters
   (City / Category / Card / date range) update every chart when you change them.

FIRST RUN
Sets itself up automatically (needs internet, about a minute). Every run after is
fast and works offline. Requires Python 3: https://www.python.org/downloads/
(On Windows, tick "Add python.exe to PATH" during install.)

MAC NOTE
The first time, macOS may block mac.command because it came from a zip.
Right-click mac.command -> Open -> Open. You only do this once.
When it finishes, press any key and the window closes. macOS may ask "Terminal
wants to control Terminal" once - click OK (or close the window yourself with Cmd+W).

BEFORE RE-RUNNING
Close expense_tracker.xlsx in Excel first. If it is open, the tracker saves a
timestamped copy instead so nothing is lost.

YOUR CATEGORIES
Your merchant-to-category choices live in app/merchant_map.json, created
automatically. To start categorizing fresh, delete that file.
