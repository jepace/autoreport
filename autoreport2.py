#!/usr/bin/env python3
'''
autoreport2.py -- next-generation version of autoreport.py

Runs the QuickBooks "Favorite Reports", exports each to Excel and saves it as
.xlsm in an output folder.  Same keystroke-driven approach as autoreport.py
(which is left untouched), restructured to fix the issues in CODE_ISSUES.md:

* All settings live in the CONFIG section below (reports, window titles,
  menu position, default output folder, timings).
* Each report declares how many date fields it has and any extra tabs it
  needs, instead of special-casing names in the code.
* No date typing: pick a year and calendar quarter (defaults to the most
  recently finished quarter, so normally you just press Enter).
* Refuses to start if Excel is already open, so it can never save or close
  the wrong workbook.
* QuickBooks is driven by keystrokes, but Excel is driven directly through
  its own automation interface (COM, via pywin32): no Save As dialog, no
  file-type dropdown, no prompts.  Saves plain .xlsx and removes the
  "QuickBooks Export Tips" sheet QB adds.
* Waits for windows and for the saved file on disk, instead of fixed sleeps.
* Every run writes a log to <output folder>/logs/.
* Failures stop with a clear message; the console waits for Enter.
  A failed run should be redone from the start.

Requirements: Windows, QuickBooks Desktop, Excel, and
    pip install pyautogui wxPython pywin32

Emergency stop: slam the mouse into a screen corner (pyautogui FAILSAFE).
'''

import datetime as dt
import logging
import os
import sys
import time
from dataclasses import dataclass

import pyautogui
import pywintypes
import win32com.client
import win32gui
import wx

# ============================================================================
# CONFIG
# ============================================================================

@dataclass
class Report:
    name: str              # Must match the name in QB's Favorite Reports, in order
    date_fields: int = 2   # 2 = From/To, 1 = single "as of" date
    extra_tabs: int = 0    # Extra tabs needed after the last date to refresh

# Order MUST match the order in QuickBooks: Reports > Favorite Reports
REPORTS = [
    Report('P-L YTD Comp'),
    Report('Balance Sheet Detail'),
    Report('Balance Sheet Summary', date_fields=1),
    Report('AP Aging Detail', date_fields=1, extra_tabs=2),
    Report('Transaction Detail by acct'),
]

QB_TITLE = 'QuickBooks Desktop Pro 2020'   # substring of the main QB window title
EXPORT_DIALOG_TITLE = 'Send Report to Excel'

def is_excel_title(title):
    # Excel windows are titled "<workbook> - Excel". Matching the ending
    # avoids browser tabs, folders etc. that merely contain "Excel".
    return title.rstrip().endswith('Excel')

# Number of "down" presses in the Reports menu to reach "Favorite Reports"
FAVORITES_MENU_DOWNS = 4

FILE_PREFIX = 'Pembrook'
FILE_EXT = 'xlsx'          # 'xlsx' (plain data) or 'xlsm' (keeps QB's macros)
EXCEL_FORMATS = {'xlsx': 51, 'xlsm': 52}   # Excel's FileFormat codes

# Sheets QuickBooks adds that we don't want. Matched case-insensitively as a
# prefix of the sheet name. Never deletes the last remaining sheet.
# Every run logs the sheet names it saw, so check the log if this needs tweaking.
DROP_SHEETS = ('QuickBooks Export Tips',)

def default_output_dir():
    base = os.environ.get('OneDrive') or os.path.expanduser('~')
    return os.path.join(base, 'Documents', 'Pembrook', 'Finance', 'Pembrook', 'auto')

# Timeouts (seconds). "Soft" waits log a warning and carry on if the window
# is never seen -- they replace old fixed sleeps whose window titles haven't
# been confirmed on the real machine yet. Check the log after a first run.
QB_REPORT_TIMEOUT = 20     # soft: report window opens
EXPORT_DIALOG_TIMEOUT = 10 # soft: "Send Report to Excel" dialog opens
EXCEL_START_TIMEOUT = 300  # Excel can be slow to start

pyautogui.PAUSE = 0.3      # 0.1 is too fast for QB; 0.3 works
CONSOLE_DEBUG = True       # show step-by-step detail in the console (log always has it)

# ============================================================================

log = logging.getLogger('autoreport')


class AutoReportError(Exception):
    pass


def setup_logging(output_dir, year, quarter):
    log.setLevel(logging.DEBUG)
    console = logging.StreamHandler(sys.stdout)
    console.setLevel(logging.DEBUG if CONSOLE_DEBUG else logging.INFO)
    console.setFormatter(logging.Formatter('%(message)s'))
    log.addHandler(console)

    log_dir = os.path.join(output_dir, 'logs')
    os.makedirs(log_dir, exist_ok=True)
    stamp = dt.datetime.now().strftime('%Y%m%d-%H%M%S')
    path = os.path.join(log_dir, f'autoreport {year}Q{quarter} {stamp}.log')
    fh = logging.FileHandler(path, encoding='utf-8')
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(logging.Formatter('%(asctime)s %(levelname)-7s %(message)s'))
    log.addHandler(fh)
    return path


# ---------------------------------------------------------------- dates ----

def quarter_range(year, quarter):
    '''(first day, last day) of calendar quarter 1-4.'''
    start = dt.date(year, 3 * quarter - 2, 1)
    next_start = dt.date(year + 1, 1, 1) if quarter == 4 else dt.date(year, 3 * quarter + 1, 1)
    return start, next_start - dt.timedelta(days=1)


def last_completed_quarter(today=None):
    '''(year, quarter) of the most recently finished calendar quarter.'''
    today = today or dt.date.today()
    q = (today.month - 1) // 3   # 0 means the last quarter was Q4 of last year
    return (today.year - 1, 4) if q == 0 else (today.year, q)


def fmt_date(d):
    # QB accepts unpadded m/d/yyyy, matching what users typed in v1
    return f'{d.month}/{d.day}/{d.year}'


def file_date(d):
    '''YYMMDD, matching the original script's file naming.'''
    return d.strftime('%y%m%d')


# ---------------------------------------------------------------- dialog ---

class InputDialog(wx.Dialog):
    def __init__(self):
        super().__init__(None, title='QuickBooks Quarterly Reports',
                         style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        self.year = self.quarter = self.output_dir = None
        year, quarter = last_completed_quarter()

        self.year_ctrl = wx.SpinCtrl(self, min=2000, max=2100, initial=year)
        self.quarter_ctrl = wx.RadioBox(self, choices=['Q1', 'Q2', 'Q3', 'Q4'],
                                        majorDimension=4, style=wx.RA_SPECIFY_COLS)
        self.quarter_ctrl.SetSelection(quarter - 1)
        self.dir_ctrl = wx.DirPickerCtrl(self, path=default_output_dir(),
                                         style=wx.DIRP_USE_TEXTCTRL)
        self.range_lbl = wx.StaticText(self)
        self.year_ctrl.Bind(wx.EVT_SPINCTRL, self.update_range)
        self.year_ctrl.Bind(wx.EVT_TEXT, self.update_range)
        self.quarter_ctrl.Bind(wx.EVT_RADIOBOX, self.update_range)
        self.update_range()

        grid = wx.FlexGridSizer(cols=2, vgap=10, hgap=10)
        grid.AddGrowableCol(1)
        for label, ctrl in (('Year', self.year_ctrl),
                            ('Quarter', self.quarter_ctrl),
                            ('Range', self.range_lbl),
                            ('Output Folder', self.dir_ctrl)):
            grid.Add(wx.StaticText(self, label=label), 0, wx.ALIGN_CENTER_VERTICAL)
            grid.Add(ctrl, 1, wx.EXPAND)

        start_btn = wx.Button(self, wx.ID_OK, 'Start')
        quit_btn = wx.Button(self, wx.ID_CANCEL, 'Quit')
        buttons = wx.StdDialogButtonSizer()
        buttons.AddButton(start_btn)
        buttons.AddButton(quit_btn)
        buttons.Realize()

        main = wx.BoxSizer(wx.VERTICAL)
        main.Add(grid, 0, wx.EXPAND | wx.ALL, 15)
        main.Add(buttons, 0, wx.EXPAND | wx.ALL, 15)
        self.SetSizerAndFit(main)
        self.SetSize((max(self.GetSize()[0], 520), -1))

        start_btn.SetDefault()
        start_btn.Bind(wx.EVT_BUTTON, self.on_start)
        self.Centre()
        start_btn.SetFocus()   # defaults are usually right: just press Enter

    def selection(self):
        return self.year_ctrl.GetValue(), self.quarter_ctrl.GetSelection() + 1

    def update_range(self, event=None):
        start, end = quarter_range(*self.selection())
        self.range_lbl.SetLabel(f'{fmt_date(start)}  -  {fmt_date(end)}')
        if event:
            event.Skip()

    def on_start(self, event):
        year, quarter = self.selection()
        _, end = quarter_range(year, quarter)
        out = self.dir_ctrl.GetPath().strip()
        if end >= dt.date.today():
            ok = wx.MessageBox(f'{fmt_date(end)} has not passed yet. Run anyway?',
                               'Quarter not finished', wx.YES_NO | wx.ICON_QUESTION, self)
            if ok != wx.YES:
                return
        if not out:
            return wx.MessageBox('Choose an output folder.', 'Invalid input',
                                 wx.OK | wx.ICON_ERROR, self)
        self.year, self.quarter, self.output_dir = year, quarter, out
        self.EndModal(wx.ID_OK)


def get_input():
    '''Returns (year, quarter, output_dir) or None if the user quit.'''
    app = wx.App(False)  # noqa: F841  (must stay alive while dialog runs)
    dlg = InputDialog()
    try:
        if dlg.ShowModal() != wx.ID_OK:
            return None
        return dlg.year, dlg.quarter, dlg.output_dir
    finally:
        dlg.Destroy()


# --------------------------------------------------------------- windows ---

def _matcher(match):
    '''A window-title test: callable, or substring (case-insensitive).'''
    if callable(match):
        return match
    needle = match.lower()
    return lambda title: needle in title.lower()


def find_windows(match):
    test = _matcher(match)
    return [w for w in pyautogui.getAllWindows() if w.title and test(w.title)]


def find_window(match):
    windows = find_windows(match)
    if len(windows) > 1:
        log.debug(f'   note: {len(windows)} windows match, using "{windows[0].title}"')
    return windows[0] if windows else None


def wait_until(condition, timeout, what, soft=False):
    '''Poll condition() until truthy. Raise (or, if soft, warn) on timeout.'''
    start = time.time()
    while True:
        result = condition()
        if result:
            log.debug(f'   ok: {what} ({time.time() - start:.1f}s)')
            return result
        if time.time() - start > timeout:
            if soft:
                log.warning(f'   did not see {what} within {timeout}s -- continuing')
                return None
            raise AutoReportError(f'Timed out after {timeout}s waiting for {what}')
        time.sleep(0.2)


def raise_window(match, timeout=30, what=None):
    what = what or f'window "{match}"'
    window = wait_until(lambda: find_window(match), timeout, what)
    log.debug(f'-> focus "{window.title}"')
    if window.isMinimized:
        window.restore()
    try:
        window.activate()
    except Exception:  # pygetwindow can raise even when activation worked
        pass
    win32gui.SetForegroundWindow(window._hWnd)
    time.sleep(0.3)
    return window


def wait_closed(match, timeout, what):
    return wait_until(lambda: not find_windows(match), timeout, f'{what} to close', soft=True)


def qb_report_is_open(name):
    '''True if a QB report window titled with `name` exists, either as a
    child window inside QuickBooks or as a separate top-level window.'''
    needle = name.lower()
    if find_window(lambda t: needle in t.lower() and not is_excel_title(t)):
        return True
    qb = find_window(QB_TITLE)
    if not qb:
        return False
    found = []
    def check(hwnd, _):
        if needle in win32gui.GetWindowText(hwnd).lower():
            found.append(hwnd)
    try:
        win32gui.EnumChildWindows(qb._hWnd, check, None)
    except Exception:
        pass
    return bool(found)


def wait_file_stable(path, timeout=60, settle=2.0):
    '''Wait until path exists and its size stays unchanged for `settle` seconds.'''
    deadline = time.time() + timeout
    last_size, stable_since = -1, None
    while time.time() < deadline:
        if os.path.isfile(path):
            size = os.path.getsize(path)
            if size > 0 and size == last_size:
                if time.time() - stable_since >= settle:
                    return
            else:
                last_size, stable_since = size, time.time()
        time.sleep(0.5)
    raise AutoReportError(f'Saved file never appeared/settled: {path}')


# ----------------------------------------------------------------- excel ---

# Excel refuses COM calls while it is busy (e.g. QB still filling the sheet).
_EXCEL_BUSY = {-2147418111,   # RPC_E_CALL_REJECTED
               -2147417846}   # RPC_E_SERVERCALL_RETRYLATER


def excel_call(fn, timeout=60):
    '''Run fn(), retrying while Excel says it is busy.'''
    deadline = time.time() + timeout
    while True:
        try:
            return fn()
        except pywintypes.com_error as e:
            if e.hresult not in _EXCEL_BUSY or time.time() > deadline:
                raise
            time.sleep(0.5)


def running_excel():
    '''The running Excel application, or None.'''
    try:
        return win32com.client.GetActiveObject('Excel.Application')
    except pywintypes.com_error:
        return None


def wait_for_export_workbook(timeout):
    '''Wait for the workbook QuickBooks creates, and for QB to finish filling it.'''
    def opened():
        xl = running_excel()
        try:
            return xl if xl and excel_call(lambda: xl.Workbooks.Count, 5) > 0 else None
        except pywintypes.com_error:
            return None
    xl = wait_until(opened, timeout, 'Excel workbook from QuickBooks')
    wb = excel_call(lambda: xl.Workbooks(1))

    # QB writes into the workbook after it opens; wait until it stops changing.
    def snapshot():
        sheets = wb.Worksheets
        return (excel_call(lambda: xl.Ready), sheets.Count,
                tuple(sheets(i).UsedRange.Address for i in range(1, sheets.Count + 1)))
    last, since = None, time.time()
    deadline = time.time() + timeout
    while time.time() < deadline:
        snap = excel_call(snapshot)
        if snap != last:
            last, since = snap, time.time()
        elif snap[0] and time.time() - since >= 2.0:
            return xl, wb
        time.sleep(0.5)
    raise AutoReportError('Excel workbook never finished loading')


def drop_unwanted_sheets(wb):
    names = [wb.Worksheets(i).Name for i in range(1, wb.Worksheets.Count + 1)]
    log.debug(f'   sheets: {names}')
    for name in names:
        if any(name.lower().startswith(p.lower()) for p in DROP_SHEETS):
            if wb.Worksheets.Count <= 1:
                log.warning(f'   not deleting "{name}": it is the only sheet')
                continue
            excel_call(lambda: wb.Worksheets(name).Delete())
            log.debug(f'   deleted sheet "{name}"')


def save_and_close(xl, wb, target):
    xl.DisplayAlerts = False        # no "macro-free workbook?" or other prompts
    drop_unwanted_sheets(wb)
    excel_call(lambda: wb.Worksheets(1).Activate())
    excel_call(lambda: wb.SaveAs(os.path.abspath(target),
                                 FileFormat=EXCEL_FORMATS[FILE_EXT]))
    log.debug(f'   SaveAs {target}')
    excel_call(lambda: wb.Close(SaveChanges=False))
    if excel_call(lambda: xl.Workbooks.Count) == 0:
        excel_call(xl.Quit)
    del wb, xl                      # release COM references so Excel can exit


def press(*keys):
    for k in keys:
        pyautogui.press(k)
    log.debug('   keys: ' + ' '.join(keys))


def hotkey(*keys):
    pyautogui.hotkey(*keys)
    log.debug('   keys: ' + '+'.join(keys))


# ------------------------------------------------------------- the steps ---

def remove_old_outputs(base):
    '''Delete earlier outputs so a stale file can't be mistaken for success.'''
    for ext in ('xlsx', 'xlsm'):
        old = f'{base}.{ext}'
        if os.path.isfile(old):
            try:
                os.remove(old)
                log.debug(f'   removed old {old}')
            except PermissionError:
                raise AutoReportError(f'Cannot delete {old} -- is it open somewhere?')


def qb_open_report(index, report, start, end):
    '''Step 1: open the report from Favorites and enter the dates.'''
    log.debug('Step 1: open report in QuickBooks')
    raise_window(QB_TITLE, what='QuickBooks')
    press('esc')                               # close previous report / home page
    hotkey('alt', 'r')                         # Reports menu
    press(*['down'] * FAVORITES_MENU_DOWNS, 'right')
    press(*['down'] * index, 'enter')          # pick the report
    wait_until(lambda: qb_report_is_open(report.name), QB_REPORT_TIMEOUT,
               f'QB report window "{report.name}"', soft=True)
    time.sleep(0.5)                            # let the window take focus
    press('tab')                               # to first date field

    if report.date_fields == 2:
        pyautogui.write(fmt_date(start))
        log.debug(f'   typed {fmt_date(start)}')
        press('tab')
    pyautogui.write(fmt_date(end))
    log.debug(f'   typed {fmt_date(end)}')
    press('tab', *['tab'] * report.extra_tabs)  # leave field so QB refreshes
    time.sleep(1)                              # no window signal for the refresh


def qb_export():
    '''Step 2: ask QuickBooks to send the report to a new Excel workbook.'''
    log.debug('Step 2: export to Excel')
    hotkey('alt', 'x')
    wait_until(lambda: find_window(EXPORT_DIALOG_TITLE), EXPORT_DIALOG_TIMEOUT,
               f'"{EXPORT_DIALOG_TITLE}" dialog', soft=True)
    press('n')                                 # "Create new worksheet"
    time.sleep(0.5)
    press('enter')


def excel_save(target):
    '''Step 3: tidy, save and close the workbook through Excel itself.'''
    log.debug('Step 3: save in Excel')
    xl, wb = wait_for_export_workbook(EXCEL_START_TIMEOUT)
    save_and_close(xl, wb, target)
    wait_file_stable(target)
    wait_closed(is_excel_title, 30, 'Excel')   # soft: next report can still proceed


def run_report(index, report, start, end, output_dir):
    log.info(f'[{index + 1}/{len(REPORTS)}] {report.name}')
    base = os.path.join(output_dir, f'{FILE_PREFIX} {file_date(end)} {report.name}')
    target = f'{base}.{FILE_EXT}'
    remove_old_outputs(base)
    qb_open_report(index, report, start, end)
    qb_export()
    excel_save(target)
    log.info(f'   saved {target}')


# ------------------------------------------------------------------ main ---

def preflight():
    '''Refuse to run if Excel is open: the script could save/close the wrong one.'''
    open_excel = find_windows(is_excel_title)
    if open_excel:
        titles = '\n    '.join(w.title for w in open_excel)
        raise AutoReportError('Close all Excel windows before running:\n    ' + titles)
    xl = running_excel()   # Excel can also be running with no visible window
    if xl is not None and excel_call(lambda: xl.Workbooks.Count, 5) > 0:
        raise AutoReportError('Excel is running in the background with workbooks open. '
                              'Close it (or end EXCEL.EXE in Task Manager) and retry.')
    if not find_window(QB_TITLE):
        raise AutoReportError(f'QuickBooks is not running (no window matching "{QB_TITLE}")')


def main():
    user = get_input()
    if user is None:
        print('Cancelled.')
        return 0
    year, quarter, output_dir = user
    start, end = quarter_range(year, quarter)
    os.makedirs(output_dir, exist_ok=True)
    log_path = setup_logging(output_dir, year, quarter)

    log.info(f'{year} Q{quarter}: {fmt_date(start)} - {fmt_date(end)}')
    log.info(f'Output: {output_dir}')
    log.info(f'Log:    {log_path}')
    try:
        preflight()
        for i, report in enumerate(REPORTS):
            run_report(i, report, start, end, output_dir)

        # Back to the QuickBooks home page
        raise_window(QB_TITLE, what='QuickBooks')
        press('esc')
        hotkey('alt', 'c')
        press('enter')
    except AutoReportError as e:
        log.error(f'** ERROR: {e}')
        log.error('Run stopped. Redo the whole run once the problem is fixed.')
        return 1
    except pyautogui.FailSafeException:
        log.error('** Stopped by user (mouse moved to screen corner).')
        return 1
    except Exception:
        log.exception('** Unexpected error')
        return 1
    log.info('REPORTS GENERATED SUCCESSFULLY')
    return 0


if __name__ == '__main__':
    code = main()
    if code:
        input('Press Enter to close...')
    sys.exit(code)
