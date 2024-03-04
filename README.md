# BGG collection in Google Sheets

A simple Apps Script for Google Sheets to fetches your BGG collection. Wrote it because I wanted to sort my games based on some columns which were not available on BGG tables.

## How to sue

1. Create a new spreadsheet.
2. Rename your sheet from "Sheet1" to "Replaceable Game Dump" (bottom of the screen).
3. _Extensions > Apps Script_. Paste the code from `main.js` there.
4. Ctrl+F `wrygiel` and replace with your username. Otherwise you'll fetch my games, not yours.
5. Return to the spreadsheet and reload the tab.
6. The "BGG" menu item should appear after a couple of seconds.
7. Click "BGG -> Reload the sheet". It will populate your sheet with the columns and rows.
8. Add some custom formatting as you wish.
