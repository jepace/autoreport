# autoreport

Runs the quarterly QuickBooks reports, exports them to Excel, and saves them
ready to send to corporate.

There are two scripts:

| Script | Status |
|---|---|
| `autoreport.py` | The original. Used for years; left untouched. |
| `autoreport2.py` | The new version, described below. Same idea, more robust. |

## autoreport2.py

### What it does

You get **one file**: a workbook with every report on its own tab, named
`Pembrook YYQN Quarterly Reports YYMMDD.xlsx` (quarter, then the date you ran
it), e.g. `Pembrook 26Q3 Quarterly Reports 261119.xlsx`.

1. Asks which **quarter** to run (it defaults to the last finished one, so
   normally you just press Enter) and shows the exact file it will create.
2. For each report in QuickBooks' **Favorite Reports** menu: opens it, enters
   the quarter's dates, sends it to Excel, checks the dates in the report's
   header, removes the "QuickBooks Desktop Export Tips" sheet, and saves it to
   a temporary working folder.
3. Combines them into the one workbook (tabs named after the titles
   QuickBooks prints, e.g. *Summary Balance Sheet*) and saves it to the
   reports folder.
4. Deletes the working folder and returns QuickBooks to its home page.

### Files and folders

| Where | What |
|---|---|
| Reports folder (default `OneDrive\Documents\Pembrook\Finance\Pembrook`) | The combined workbook. **Never overwritten:** if the name is taken (e.g. a second run the same day), it is saved as `... v2.xlsx`, `... v3.xlsx`, etc. |
| `<reports folder>\auto\logs\` | A log of every run |
| Windows temp folder (`%TEMP%\autoreport ...`) | Temporary per-report files, kept out of OneDrive. Deleted after a successful run; **kept after a failed one** so you can see how far it got (the log gives the path) |

Nothing is put in the reports folder unless every report succeeded.

QuickBooks is driven with keystrokes. Excel is controlled directly through its
own automation interface, so there's no Save As dialog and no pop-ups.

### Requirements

* Windows 10/11, QuickBooks Desktop Pro 2020, Microsoft 365 Excel
* Python 3.8+ (e.g. from the Microsoft Store)
* `pip install pyautogui wxPython pywin32`

### One-time setup

Add the reports to **Reports > Favorite Reports** in QuickBooks, in this order:

1. P-L YTD Comp
2. Balance Sheet Detail
3. Balance Sheet Summary
4. AP Aging Detail
5. Transaction Detail by acct

The script picks reports by their **position** in that menu, so the order has
to match the `REPORTS` list at the top of `autoreport2.py`.

### Running it

1. Open QuickBooks. **Close Excel**: the script refuses to start if Excel is
   open, so it can't save or close the wrong workbook.
2. Run `python autoreport2.py`.
3. Check the quarter and the file name it shows, then press **Start** (or
   Enter). Use **Change…** to put the file somewhere else.
4. Don't touch the keyboard or mouse until it finishes. **To stop it in an
   emergency, move the mouse hard into a corner of the screen.**

If a run fails, fix the problem and **redo the whole run**. Close any Excel
window it left open first (the script will remind you). You can delete the
leftover working folder (its path is at the end of the log) once you've looked
at it.

### Log files

Every run writes a log to `<reports folder>\auto\logs\`, named like
`autoreport 2026Q3 20261119-191736.log`. It records each keystroke, each
window it waited for, the sheet names in each workbook, and every file saved.
Check it first if something goes wrong.

Two things in the log deserve attention:

* `** <report>: expected "..." in the page header -- CHECK THE DATES`
  means the exported report may have the wrong date range. Open that tab and
  check before sending it.
* `The "<title>" report window never appeared` stops the run. Either the wrong
  report opened (check the Favorites order), or the report's `title` in
  `REPORTS` doesn't match. The log lists the visible QuickBooks windows just
  above, so you can copy the right title from there.

### Settings

All settings are at the top of `autoreport2.py`:

| Setting | What it controls |
|---|---|
| `REPORTS` | Reports in Favorites order: the Favorites name (for reference), the title QuickBooks prints on the report (used to find its window, and as the tab name), how many date fields it has (`date_fields`), extra tabs needed to refresh (`extra_tabs`), whether its header reads "As of" (`as_of`), and an optional different tab name (`tab`) |
| `QB_TITLE` | Part of the main QuickBooks window title |
| `FAVORITES_MENU_DOWNS` | How far down the Reports menu "Favorite Reports" is |
| `FILE_PREFIX` | First word of each file name |
| `FILE_EXT` | `'xlsx'` (plain data) or `'xlsm'` (keeps QuickBooks' macros) |
| `EXPORT_TIPS_SHEET` | The tips sheet QuickBooks adds when it *finishes* exporting. The script waits for it, then deletes it |
| `combined_name()` | Name of the combined workbook |
| `default_reports_dir()` | Default reports folder (under OneDrive) |
| `WORK_SUBDIR` | Subfolder of the reports folder that holds `logs\` (`auto`) |
| `QUARTERS_LISTED` | How many quarters the dialog offers (8) |
| `*_TIMEOUT` | How long to wait for each window |
| `pyautogui.PAUSE` | Delay between keystrokes (0.3 s; 0.1 is too fast for QuickBooks) |
| `CONSOLE_DEBUG` | Step-by-step detail in the console (the log always has it) |

### Known limits

* Reports are chosen by menu position, so reordering Favorites or a QuickBooks
  update that changes the Reports menu will pick the wrong report. The
  script notices (the report window it expects never appears) and stops.
* QuickBooks itself is still driven by keystrokes, so leave the PC alone
  while it runs.

### If it breaks after a QuickBooks or Excel update

The script stops rather than guessing, and the log says where. The usual
suspects, all in the settings at the top of `autoreport2.py`:

| Log says | Likely fix |
|---|---|
| `report window never appeared` | Favorites order changed, or QuickBooks renamed the report: update `REPORTS` (the log lists the window titles it saw) |
| `did not open the Excel menu` | QuickBooks was slow or the toolbar changed; try once more, then check Alt+X still opens the report's Excel button menu |
| `QuickBooks is not running` | The main window title changed: update `QB_TITLE` |
| `did not finish exporting` | The tips sheet was renamed or turned off: update `EXPORT_TIPS_SHEET` |
