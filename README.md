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

1. Asks for the **year and calendar quarter** (it defaults to the most
   recently finished quarter, so normally you just press Enter) and the output
   folder.
2. For each report in QuickBooks' **Favorite Reports** menu:
   * opens it, enters the quarter's dates, and sends it to Excel;
   * removes the "QuickBooks Export Tips" sheet;
   * saves it as `Pembrook YYMMDD <report name>.xlsx`, then closes Excel.
3. Builds one extra workbook, `Pembrook YYMMDD All Reports.xlsx`, with every
   report as its own tab.
4. Returns QuickBooks to its home page.

`YYMMDD` is the last day of the quarter, e.g. `260930` for 2026 Q3.

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
3. Check the quarter and folder, then press **Start**.
4. Don't touch the keyboard or mouse until it finishes. **To stop it in an
   emergency, move the mouse hard into a corner of the screen.**

If a run fails, fix the problem and **redo the whole run**. The script
deletes old copies of each file before recreating it, so nothing stale is left
behind.

### Log files

Every run writes a log to `<output folder>\logs\`, named like
`autoreport 2026Q3 20261119-191736.log`. It records each keystroke, each
window it waited for, the sheet names in each workbook, and every file saved.
Check it first if something goes wrong.

Warnings like `did not see QB report window ... -- continuing` mean a window
title in the settings doesn't match what QuickBooks shows. The run still
works, just more slowly. Correct the title in the settings to fix it.

### Settings

All settings are at the top of `autoreport2.py`:

| Setting | What it controls |
|---|---|
| `REPORTS` | Reports in Favorites order; how many date fields each has (`date_fields`); extra tabs needed to refresh (`extra_tabs`) |
| `QB_TITLE`, `EXPORT_DIALOG_TITLE` | Window titles the script looks for |
| `FAVORITES_MENU_DOWNS` | How far down the Reports menu "Favorite Reports" is |
| `FILE_PREFIX` | First word of each file name |
| `FILE_EXT` | `'xlsx'` (plain data) or `'xlsm'` (keeps QuickBooks' macros) |
| `DROP_SHEETS` | Sheet names to delete from each export |
| `COMBINED_WORKBOOK`, `COMBINED_NAME` | Whether to build the all-in-one workbook, and its name |
| `default_output_dir()` | Default output folder (under OneDrive) |
| `*_TIMEOUT` | How long to wait for each window |
| `pyautogui.PAUSE` | Delay between keystrokes (0.3 s; 0.1 is too fast for QuickBooks) |
| `CONSOLE_DEBUG` | Step-by-step detail in the console (the log always has it) |

### Known limits

* Reports are chosen by menu position, so reordering Favorites or a QuickBooks
  update that changes the Reports menu will pick the wrong report.
* QuickBooks itself is still driven by keystrokes, so leave the PC alone
  while it runs.
