#!/usr/bin/env python3
'''
autoreport2.py -- next-generation version of autoreport.py

Runs the QuickBooks "Favorite Reports", exports each to Excel, and combines
them into ONE workbook (one tab per report) in the reports folder:

    Pembrook 26Q3 Quarterly Reports 261119.xlsx      (quarter, then run date)

Same keystroke-driven approach as autoreport.py (which is left untouched),
restructured to fix the issues in CODE_ISSUES.md:

* All settings live in the CONFIG section below.
* Pick the quarter from a list (defaults to the last finished one, so
  normally you just press Enter). No date typing.
* Refuses to start if Excel is already open, so it can never save or close
  the wrong workbook.
* QuickBooks is driven by keystrokes; Excel is driven directly through its
  own automation interface (COM, via pywin32), with no dialogs or prompts.
  Plain .xlsx, with QuickBooks' "Export Tips" sheet removed.
* Never overwrites: if the name is taken, saves as "... v2.xlsx", etc.
* The per-report files are made in a temporary working folder under
  <reports folder>\\auto and deleted after a successful run, so the only
  file you see is the combined one.  A failed run keeps them for inspection.
* Waits for windows and files instead of fixed sleeps.
* Every run writes a log to <reports folder>\\auto\\logs.
* A failed run should be redone from the start.

Requirements: Windows, QuickBooks Desktop, Excel, and
    pip install pyautogui wxPython pywin32

Emergency stop: slam the mouse into a screen corner (pyautogui FAILSAFE).
'''

import ctypes
import datetime as dt
import logging
import os
import shutil
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

# Order MUST match the order in QuickBooks: Reports > Favorite Reports.
# The combined workbook's tabs are in this order too.
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

def combined_name(year, quarter, run_date):
    '''e.g. "Pembrook 26Q3 Quarterly Reports 261119" (run date at the end).'''
    return f'{FILE_PREFIX} {year % 100:02d}Q{quarter} Quarterly Reports {run_date:%y%m%d}'

# Sheets QuickBooks adds that we don't want. Matched case-insensitively as a
# prefix of the sheet name. Never deletes the last remaining sheet.
# Every run logs the sheet names it saw, so check the log if this needs tweaking.
DROP_SHEETS = ('QuickBooks Export Tips',)

def default_reports_dir():
    '''Where the combined workbook goes (can be changed in the dialog).'''
    base = os.environ.get('OneDrive') or os.path.expanduser('~')
    return os.path.join(base, 'Documents', 'Pembrook', 'Finance', 'Pembrook')

WORK_SUBDIR = 'auto'       # <reports folder>\auto holds logs and working files
QUARTERS_LISTED = 8        # how many quarters the dialog offers

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


def setup_logging(log_dir, year, quarter, stamp):
    log.setLevel(logging.DEBUG)
    console = logging.StreamHandler(sys.stdout)
    console.setLevel(logging.DEBUG if CONSOLE_DEBUG else logging.INFO)
    console.setFormatter(logging.Formatter('%(message)s'))
    log.addHandler(console)

    os.makedirs(log_dir, exist_ok=True)
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


def previous_quarter(year, quarter):
    return (year - 1, 4) if quarter == 1 else (year, quarter - 1)


def recent_quarters(count, today=None):
    '''[(year, quarter), ...] newest first, starting with the one in progress.'''
    today = today or dt.date.today()
    yq = (today.year, (today.month - 1) // 3 + 1)
    result = []
    for _ in range(count):
        result.append(yq)
        yq = previous_quarter(*yq)
    return result


def fmt_date(d):
    # QB accepts unpadded m/d/yyyy, matching what users typed in v1
    return f'{d.month}/{d.day}/{d.year}'


def file_date(d):
    '''YYMMDD, matching the original script's file naming.'''
    return d.strftime('%y%m%d')


# ---------------------------------------------------------------- dialog ---

def enable_high_dpi():
    '''Tell Windows we handle high-DPI screens, so text is sharp, not stretched.'''
    if sys.platform != 'win32':
        return
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)   # per-monitor (Win 8.1+)
    except (AttributeError, OSError):
        try:
            ctypes.windll.user32.SetProcessDPIAware()    # older fallback
        except (AttributeError, OSError):
            pass


class InputDialog(wx.Dialog):
    '''Asks which quarter to run and where the combined workbook goes.

    After ShowModal() returns wx.ID_OK, self.result is
    (year, quarter, reports_dir); otherwise it is None.'''

    def __init__(self, reports_dir, today=None):
        super().__init__(None, title='Quarterly Reports')
        self.today = today or dt.date.today()
        self.reports_dir = reports_dir
        self.quarters = recent_quarters(QUARTERS_LISTED, self.today)
        self.result = None

        self._build()
        self._update_preview()
        self.Centre()

    # -- layout --

    def _build(self):
        gap = self.FromDIP(8)
        margin = self.FromDIP(20)

        title = wx.StaticText(self, label='QuickBooks Quarterly Reports')
        title.SetFont(title.GetFont().Scaled(1.4).Bold())

        self.quarter_choice = wx.Choice(self, choices=[self._quarter_label(*yq)
                                                       for yq in self.quarters])
        self.quarter_choice.SetSelection(1)   # [0] is in progress; [1] last finished
        self.quarter_choice.Bind(wx.EVT_CHOICE, self._update_preview)

        self.file_text = wx.StaticText(self)
        self.file_text.SetFont(self.file_text.GetFont().Bold())
        self.dir_text = wx.StaticText(self, style=wx.ST_ELLIPSIZE_MIDDLE)
        change = wx.Button(self, label='Change…')
        change.Bind(wx.EVT_BUTTON, self._on_change_dir)

        dir_row = wx.BoxSizer(wx.HORIZONTAL)
        dir_row.Add(self.dir_text, 1, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, gap)
        dir_row.Add(change, 0, wx.ALIGN_CENTER_VERTICAL)

        form = wx.FlexGridSizer(cols=2, vgap=gap, hgap=self.FromDIP(12))
        form.AddGrowableCol(1)
        for label, item in (('Quarter', self.quarter_choice),
                            ('Creates', self.file_text),
                            ('In', dir_row)):
            form.Add(wx.StaticText(self, label=label), 0, wx.ALIGN_CENTER_VERTICAL)
            form.Add(item, 1, wx.EXPAND | wx.ALIGN_CENTER_VERTICAL)

        notice = self._notice('Close Excel first. Don\'t touch the keyboard or mouse '
                              'while it runs. To stop it, move the mouse into a '
                              'corner of the screen.')

        buttons = self.CreateStdDialogButtonSizer(wx.OK | wx.CANCEL)
        self.FindWindow(wx.ID_OK).SetLabel('Start')
        self.FindWindow(wx.ID_CANCEL).SetLabel('Quit')
        self.Bind(wx.EVT_BUTTON, self._on_start, id=wx.ID_OK)

        body = wx.BoxSizer(wx.VERTICAL)
        body.Add(title, 0, wx.BOTTOM, margin)
        body.Add(form, 0, wx.EXPAND | wx.BOTTOM, margin)
        body.Add(notice, 0, wx.EXPAND | wx.BOTTOM, margin)
        body.Add(buttons, 0, wx.ALIGN_RIGHT)

        outer = wx.BoxSizer(wx.VERTICAL)
        outer.Add(body, 1, wx.EXPAND | wx.ALL, margin)
        outer.SetMinSize(self.FromDIP(wx.Size(560, -1)))
        self.SetSizerAndFit(outer)
        self.FindWindow(wx.ID_OK).SetFocus()   # Enter = Start

    def _notice(self, message):
        '''An info icon beside wrapped text.'''
        icon = wx.StaticBitmap(self, bitmap=wx.ArtProvider.GetBitmap(
            wx.ART_INFORMATION, wx.ART_MESSAGE_BOX, self.FromDIP(wx.Size(20, 20))))
        text = wx.StaticText(self, label=message)
        text.Wrap(self.FromDIP(470))
        row = wx.BoxSizer(wx.HORIZONTAL)
        row.Add(icon, 0, wx.RIGHT, self.FromDIP(10))
        row.Add(text, 1)
        return row

    # -- behaviour --

    def _quarter_label(self, year, quarter):
        start, end = quarter_range(year, quarter)
        label = f'{year} Q{quarter}   ({start:%b} {start.day} – {end:%b} {end.day})'
        return label + '   — in progress' if end >= self.today else label

    def _selected(self):
        return self.quarters[self.quarter_choice.GetSelection()]

    def _update_preview(self, event=None):
        self.file_text.SetLabel(f'{combined_name(*self._selected(), self.today)}.{FILE_EXT}')
        self.dir_text.SetLabel(self.reports_dir)
        self.dir_text.SetToolTip(self.reports_dir)
        self.Layout()

    def _on_change_dir(self, event):
        with wx.DirDialog(self, 'Where should the combined workbook go?',
                          defaultPath=self.reports_dir,
                          style=wx.DD_DEFAULT_STYLE | wx.DD_DIR_MUST_EXIST) as dlg:
            if dlg.ShowModal() == wx.ID_OK:
                self.reports_dir = dlg.GetPath()
                self._update_preview()

    def _on_start(self, event):
        year, quarter = self._selected()
        _, end = quarter_range(year, quarter)
        if end >= self.today and wx.MessageBox(
                f'{year} Q{quarter} isn\'t over until {fmt_date(end)}. Run anyway?',
                'Quarter in progress', wx.YES_NO | wx.NO_DEFAULT | wx.ICON_QUESTION,
                self) != wx.YES:
            return
        if not os.path.isdir(self.reports_dir):
            wx.MessageBox(f'Folder not found:\n{self.reports_dir}', 'Folder not found',
                          wx.OK | wx.ICON_ERROR, self)
            return
        self.result = (year, quarter, self.reports_dir)
        self.EndModal(wx.ID_OK)


def get_input():
    '''Show the dialog. Returns (year, quarter, reports_dir), or None if the user quit.'''
    enable_high_dpi()
    app = wx.App(False)  # noqa: F841  (must exist while the dialog runs)
    dlg = InputDialog(default_reports_dir())
    try:
        dlg.ShowModal()
        return dlg.result
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


def press(*keys):
    for k in keys:
        pyautogui.press(k)
    log.debug('   keys: ' + ' '.join(keys))


def hotkey(*keys):
    pyautogui.hotkey(*keys)
    log.debug('   keys: ' + '+'.join(keys))


# ----------------------------------------------------------------- files ---

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


def publish(src, dest_dir):
    '''Copy src into dest_dir without ever overwriting. Returns the new path.'''
    stem, ext = os.path.splitext(os.path.basename(src))
    n = 1
    while True:
        name = f'{stem}{ext}' if n == 1 else f'{stem} v{n}{ext}'
        dest = os.path.join(dest_dir, name)
        try:
            with open(src, 'rb') as fin, open(dest, 'xb') as fout:  # 'x' fails if it exists
                shutil.copyfileobj(fin, fout)
        except FileExistsError:
            log.debug(f'   {name} already exists in {dest_dir}')
            n += 1
            continue
        shutil.copystat(src, dest)
        return dest


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


def sheet_name(text):
    '''Make text a legal Excel tab name (31 chars, no []:*?/\\).'''
    for ch in '[]:*?/\\':
        text = text.replace(ch, '-')
    return text[:31]


def build_combined(saved, target):
    '''Copy the sheets of each saved report into one workbook, one tab per report.

    saved: list of (report name, file path), in tab order.
    Uses its own hidden Excel instance and always quits it.'''
    log.info(f'Combining {len(saved)} reports')
    xl = win32com.client.DispatchEx('Excel.Application')
    try:
        xl.Visible = False
        xl.DisplayAlerts = False
        combo = excel_call(lambda: xl.Workbooks.Add())
        blanks = [combo.Worksheets(i).Name for i in range(1, combo.Worksheets.Count + 1)]
        for name, path in saved:
            src = excel_call(lambda: xl.Workbooks.Open(os.path.abspath(path), ReadOnly=True))
            count = src.Worksheets.Count
            for i in range(1, count + 1):
                last = combo.Worksheets(combo.Worksheets.Count)
                excel_call(lambda: src.Worksheets(i).Copy(Before=None, After=last))
                copied = combo.Worksheets(combo.Worksheets.Count)
                copied.Name = sheet_name(name if count == 1 else f'{name} {i}')
                log.debug(f'   tab "{copied.Name}" <- {os.path.basename(path)}')
            excel_call(lambda: src.Close(SaveChanges=False))
        for blank in blanks:                       # the empty sheets Add() made
            excel_call(lambda: combo.Worksheets(blank).Delete())
        excel_call(lambda: combo.Worksheets(1).Activate())
        excel_call(lambda: combo.SaveAs(os.path.abspath(target),
                                        FileFormat=EXCEL_FORMATS[FILE_EXT]))
        excel_call(lambda: combo.Close(SaveChanges=False))
    finally:
        try:
            xl.Quit()
        except pywintypes.com_error:
            pass
        del xl
    wait_file_stable(target)


# ------------------------------------------------------------- the steps ---

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


def run_report(index, report, start, end, work_dir):
    '''Produce one report's workbook in work_dir; returns its path.'''
    log.info(f'[{index + 1}/{len(REPORTS)}] {report.name}')
    target = os.path.join(work_dir, f'{FILE_PREFIX} {file_date(end)} {report.name}.{FILE_EXT}')
    qb_open_report(index, report, start, end)
    qb_export()
    excel_save(target)
    log.debug(f'   saved {target}')
    return target


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


def run(year, quarter, reports_dir, work_dir):
    '''Make every report, combine them, and publish. Returns the published path.'''
    start, end = quarter_range(year, quarter)
    preflight()
    saved = [(report.name, run_report(i, report, start, end, work_dir))
             for i, report in enumerate(REPORTS)]
    combined = os.path.join(work_dir,
                            f'{combined_name(year, quarter, dt.date.today())}.{FILE_EXT}')
    build_combined(saved, combined)
    published = publish(combined, reports_dir)

    # Back to the QuickBooks home page
    raise_window(QB_TITLE, what='QuickBooks')
    press('esc')
    hotkey('alt', 'c')
    press('enter')
    return published


def main():
    choice = get_input()
    if choice is None:
        print('Cancelled.')
        return 0
    year, quarter, reports_dir = choice
    stamp = dt.datetime.now().strftime('%Y%m%d-%H%M%S')
    auto_dir = os.path.join(reports_dir, WORK_SUBDIR)
    work_dir = os.path.join(auto_dir, f'working {year}Q{quarter} {stamp}')
    os.makedirs(work_dir)
    log_path = setup_logging(os.path.join(auto_dir, 'logs'), year, quarter, stamp)

    start, end = quarter_range(year, quarter)
    log.info(f'{year} Q{quarter}: {fmt_date(start)} - {fmt_date(end)}')
    log.info(f'Reports folder: {reports_dir}')
    log.info(f'Log: {log_path}')
    try:
        published = run(year, quarter, reports_dir, work_dir)
    except AutoReportError as e:
        log.error(f'** ERROR: {e}')
    except pyautogui.FailSafeException:
        log.error('** Stopped by user (mouse moved to screen corner).')
    except Exception:
        log.exception('** Unexpected error')
    else:
        shutil.rmtree(work_dir, ignore_errors=True)
        log.info(f'\nDONE: {published}')
        return 0
    log.error(f'Run stopped. Working files kept in: {work_dir}')
    log.error('Redo the whole run once the problem is fixed.')
    return 1


if __name__ == '__main__':
    code = main()
    if code:
        input('Press Enter to close...')
    sys.exit(code)
