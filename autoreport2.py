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
* The per-report files are made in a temporary working folder (in the local
  temp folder, outside OneDrive) and deleted after a successful run, so the
  only file you see is the combined one.  A failed run keeps them, and the
  log says where.
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
import re
import shutil
import sys
import tempfile
import threading
import time
from dataclasses import dataclass

import pyautogui
import pythoncom
import pywintypes
import win32com.client
import win32gui
import win32process
import wx

# ============================================================================
# CONFIG
# ============================================================================

@dataclass
class Report:
    name: str              # Name in QB's Favorite Reports (for reference only)
    title: str             # Title QB prints on the report and gives its window
    date_fields: int = 2   # 2 = From/To, 1 = single "as of" date
    extra_tabs: int = 0    # Extra tabs needed after the last date to refresh
    as_of: bool = None     # Header reads "As of <end>" (default: if 1 date field)
    tab: str = None        # Tab name in the combined workbook (default: title)

    def __post_init__(self):
        if self.as_of is None:
            self.as_of = self.date_fields == 1
        self.tab = self.tab or self.title

# Order MUST match the order in QuickBooks: Reports > Favorite Reports.
# The combined workbook's tabs are in this order too, named after QB's titles.
REPORTS = [
    Report('P-L YTD Comp', 'Profit & Loss YTD Comparison'),
    # Point-in-time report: QB headers it "As of <end>" despite From/To fields
    Report('Balance Sheet Detail', 'Balance Sheet Detail', as_of=True),
    Report('Balance Sheet Summary', 'Summary Balance Sheet', date_fields=1),
    Report('AP Aging Detail', 'A/P Aging Detail', date_fields=1, extra_tabs=2,
           tab='AP Aging Detail'),              # Excel tab names can't contain "/"
    Report('Transaction Detail by acct', 'Transaction Detail by Account'),
]

QB_TITLE = 'QuickBooks Desktop Pro 2020'   # substring of the main QB window title

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

# QuickBooks adds a "QuickBooks Desktop Export Tips" sheet when it FINISHES
# exporting, so its appearance is the signal that the export is complete.
# It is then deleted. Matched case-insensitively anywhere in the sheet name.
EXPORT_TIPS_SHEET = 'Export Tips'

def default_reports_dir():
    '''Where the combined workbook goes (can be changed in the dialog).'''
    base = os.environ.get('OneDrive') or os.path.expanduser('~')
    return os.path.join(base, 'Documents', 'Pembrook', 'Finance', 'Pembrook')

WORK_SUBDIR = 'auto'       # <reports folder>\auto\logs holds the logs
QUARTERS_LISTED = 8        # how many quarters the dialog offers

# Waits (seconds).
QB_REPORT_TIMEOUT = 60     # for the report window to appear (big reports are slow)
QB_REPORT_SETTLE = 1.5     # extra pause after it appears, before typing dates
EXPORT_TIMEOUT = 900       # for QB to finish exporting to Excel (big reports: minutes)

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


def _bring_to_front(window):
    window.Raise()
    try:
        win32gui.SetForegroundWindow(window.GetHandle())
    except Exception:   # Windows may refuse; the dialog is still shown
        pass


def get_input():
    '''Show the dialog. Returns (year, quarter, reports_dir), or None if the user quit.'''
    enable_high_dpi()
    app = wx.App(False)  # noqa: F841  (must exist while the dialog runs)
    dlg = InputDialog(default_reports_dir())
    try:
        # Launched from a terminal, Windows may open the dialog behind it.
        wx.CallAfter(_bring_to_front, dlg)
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


def wait_until(condition, timeout, what, soft=False, minimum=0, on_miss=None):
    '''Poll condition() until truthy. Raise (or, if soft, warn) on timeout.

    minimum: always take at least this long, even if condition is met sooner.
    on_miss: called on a soft timeout, e.g. to log diagnostics.'''
    start = time.time()
    while True:
        result = condition()
        if result:
            log.debug(f'   ok: {what} ({time.time() - start:.1f}s)')
            time.sleep(max(0.0, minimum - (time.time() - start)))
            return result
        if time.time() - start > minimum + timeout:
            if soft:
                log.warning(f'   did not see {what} -- continuing')
                if on_miss:
                    on_miss()
                return None
            raise AutoReportError(f'Timed out after {timeout}s waiting for {what}')
        time.sleep(0.2)


def qb_dialogs():
    '''{hwnd: title} of visible top-level windows owned by QuickBooks other
    than its main window -- i.e. its dialogs.'''
    qb = find_window(QB_TITLE)
    if not qb:
        return {}
    _, pid = win32process.GetWindowThreadProcessId(qb._hWnd)
    found = {}
    def collect(hwnd, _):
        if (hwnd != qb._hWnd and win32gui.IsWindowVisible(hwnd)
                and win32process.GetWindowThreadProcessId(hwnd)[1] == pid):
            found[hwnd] = win32gui.GetWindowText(hwnd)
    try:
        win32gui.EnumWindows(collect, None)
    except Exception:
        pass
    return found


QB_MENU_CLASS = 'XTPSkinManagerMenu'   # QuickBooks' dropdown-menu windows


def new_qb_windows(before, menus):
    '''{hwnd: "title [class]"} of QB windows not in `before`: only dropdown
    menus if `menus`, otherwise only non-menu windows (dialogs).'''
    found = {}
    for hwnd, title in qb_dialogs().items():
        if hwnd in before:
            continue
        cls = win32gui.GetClassName(hwnd)
        if (cls == QB_MENU_CLASS) == menus:
            found[hwnd] = f'{title!r} [{cls}]'
    return found


def log_qb_windows():
    '''Diagnostics: log the visible windows inside QuickBooks.'''
    titles = list(dict.fromkeys(text for _, text in qb_inner_windows()))
    log.warning('   visible QuickBooks windows: ' + ' | '.join(titles[:40]))


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


def qb_inner_windows(visible_only=True):
    '''[(hwnd, title)] of the windows inside QuickBooks.

    QB keeps closed forms around as hidden windows, so by default only
    visible ones count -- otherwise a stale window can look like the report.'''
    qb = find_window(QB_TITLE)
    if not qb:
        return []
    found = []
    def collect(hwnd, _):
        text = win32gui.GetWindowText(hwnd).strip()
        if text and (not visible_only or win32gui.IsWindowVisible(hwnd)):
            found.append((hwnd, text))
    try:
        win32gui.EnumChildWindows(qb._hWnd, collect, None)
    except Exception:
        pass
    return found


def qb_report_window(title):
    '''hwnd of the visible QB report window with exactly this title, or None.'''
    for hwnd, text in qb_inner_windows():
        if text.lower() == title.lower():
            return hwnd
    return None


def press(*keys):
    for k in keys:
        pyautogui.press(k)
    log.debug('   keys: ' + ' '.join(keys))


def hotkey(*keys):
    pyautogui.hotkey(*keys)
    log.debug('   keys: ' + '+'.join(keys))


# ----------------------------------------------------------------- files ---

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


def _export_snapshot():
    '''(xl, wb, state) for the newest open workbook, or None if there isn't one.

    Looked up fresh every time: while exporting, QuickBooks can replace the
    workbook (or reconnect to Excel), which "disconnects" any reference we
    held. state captures everything that changes while QB is still writing.'''
    xl = running_excel()
    if xl is None:
        return None
    try:
        count = xl.Workbooks.Count
        if count == 0:
            return None
        wb = xl.Workbooks(count)
        sheets = wb.Worksheets
        names = tuple(sheets(i).Name for i in range(1, sheets.Count + 1))
        state = (xl.Ready, wb.FullName, names,
                 tuple(sheets(i).UsedRange.Address for i in range(1, sheets.Count + 1)))
        return xl, wb, state
    except pywintypes.com_error as e:
        log.debug(f'   Excel not ready ({e.args[1] if len(e.args) > 1 else e})')
        return None


def is_tips_sheet(name):
    return EXPORT_TIPS_SHEET.lower() in name.lower()


def qb_busy():
    '''True while QuickBooks is showing a modal window (e.g. export progress).'''
    qb = find_window(QB_TITLE)
    try:
        return bool(qb) and not win32gui.IsWindowEnabled(qb._hWnd)
    except Exception:
        return False


def wait_for_export_workbook(timeout, settle=3.0):
    '''Wait for the workbook QuickBooks creates, and for QB to FINISH filling it.

    Done means: QB has added its Export Tips sheet (it does that last), QB is
    no longer busy, Excel is ready, and nothing has changed for `settle` s.
    Quiet alone isn't enough: on big reports QB can go quiet for a long time
    mid-export.'''
    start = time.time()
    last_state, since, seen, last_note = None, None, False, 0
    while time.time() - start < timeout:
        snap = _export_snapshot()
        if snap is None:
            last_state = None                  # gone or busy: start settling over
        else:
            xl, wb, state = snap
            ready, _, names, _ = state
            if not seen:
                log.debug(f'   Excel workbook appeared ({time.time() - start:.1f}s)')
                seen = True
            finished = any(is_tips_sheet(n) for n in names) and not qb_busy()
            if state != last_state:
                last_state, since = state, time.time()
            elif ready and finished and time.time() - since >= settle:
                log.debug(f'   export finished: {list(names)} ({time.time() - start:.1f}s)')
                return xl, wb
            if time.time() - last_note > 30:   # reassurance on long exports
                log.info(f'   waiting for QuickBooks to finish exporting '
                         f'({time.time() - start:.0f}s)...')
                last_note = time.time()
        time.sleep(0.5)
    raise AutoReportError(
        f'QuickBooks did not finish exporting to Excel within {timeout}s'
        if seen else 'Excel workbook from QuickBooks never appeared')


def drop_tips_sheets(wb):
    for name in [wb.Worksheets(i).Name for i in range(1, wb.Worksheets.Count + 1)]:
        if is_tips_sheet(name) and wb.Worksheets.Count > 1:
            excel_call(lambda: wb.Worksheets(name).Delete())
            log.debug(f'   deleted sheet "{name}"')


_MONTHS = ('January February March April May June July August '
           'September October November December').split()


def page_header_text(ws):
    '''The report's page header (company, title, date range) as plain text.
    QB puts it in Excel's Page Setup header; strip Excel's formatting codes
    like &"Arial,Bold"&14.'''
    setup = ws.PageSetup
    raw = ' '.join(h for h in (setup.LeftHeader, setup.CenterHeader, setup.RightHeader) if h)
    text = re.sub(r'&"[^"]*"|&\d+|&[A-Za-z]', ' ', raw)
    return ' '.join(text.split())


def check_report_dates(wb, report, start, end):
    '''Check the report's date range in its page header, in QB's own wording:
    "As of September 30, 2026" (one date) or "July through September 2026" /
    "July 1 through September 30, 2026" (two dates). Only the page header is
    checked: data rows contain transaction dates that would match by chance.
    Logs a loud WARNING on a mismatch; does not stop the run.'''
    try:
        ws = next(wb.Worksheets(i) for i in range(1, wb.Worksheets.Count + 1)
                  if not is_tips_sheet(wb.Worksheets(i).Name))
        header = excel_call(lambda: page_header_text(ws))
    except (pywintypes.com_error, StopIteration) as e:
        log.warning(f'   ** {report.title}: could not read the page header to check dates ({e})')
        return
    s_month, e_month = _MONTHS[start.month - 1], _MONTHS[end.month - 1]
    if report.as_of:
        expected = f'As of {e_month} {end.day}, {end.year}'
        ok = expected in header
    else:
        expected = f'{s_month} through {e_month} {end.year}'
        ok = bool(re.search(rf'\b{s_month}( {start.day})? through {e_month}'
                            rf'( {end.day},)? {end.year}\b', header))
    if ok:
        log.debug(f'   dates ok: {header}')
    else:
        log.warning(f'   ** {report.title}: expected "{expected}" in the page header '
                    f'-- CHECK THE DATES. Header: {header or "(empty)"}')


def save_and_close(xl, wb, target):
    xl.DisplayAlerts = False        # no "macro-free workbook?" or other prompts
    drop_tips_sheets(wb)
    excel_call(lambda: wb.Worksheets(1).Activate())
    excel_call(lambda: wb.SaveAs(os.path.abspath(target),
                                 FileFormat=EXCEL_FORMATS[FILE_EXT]))
    log.debug(f'   SaveAs {target}')
    excel_call(lambda: wb.Close(SaveChanges=False))


_excel_closers = []   # background threads shutting Excel instances down


def excel_closer(xl):
    '''A not-yet-started thread that quits this Excel in the background.

    Excel takes up to ~15s to shut down, and dropping our last link to a
    quitting Excel blocks until it's gone -- that was a 10-15s pause per
    report. Instead a helper thread quits Excel and does the waiting, while
    the run carries on. Usage (order matters):

        closer = excel_closer(xl)
        del xl, wb        # drop our own links while Excel is still running: instant
        closer.start()'''
    stream = pythoncom.CoMarshalInterThreadInterfaceInStream(
        pythoncom.IID_IDispatch, xl._oleobj_)

    def worker():
        pythoncom.CoInitialize()
        try:
            app = win32com.client.Dispatch(pythoncom.CoGetInterfaceAndReleaseStream(
                stream, pythoncom.IID_IDispatch))
            t = time.time()
            if app.Workbooks.Count == 0:   # never quit an Excel that has work open
                app.Quit()
            del app
            log.debug(f'   (background) Excel closed ({time.time() - t:.1f}s)')
        except Exception as e:
            log.debug(f'   (background) closing Excel: {e}')
        finally:
            pythoncom.CoUninitialize()

    thread = threading.Thread(target=worker, daemon=True)
    _excel_closers.append(thread)
    return thread


def wait_for_excel_closers(timeout=60):
    '''At the end of the run, let background Excel shutdowns finish.'''
    for thread in _excel_closers:
        thread.join(timeout)


def check_saved(path):
    '''SaveAs is synchronous, so the file must already be there.'''
    if not (os.path.isfile(path) and os.path.getsize(path) > 0):
        raise AutoReportError(f'Saved file is missing or empty: {path}')


def sheet_name(text):
    '''Make text a legal Excel tab name (31 chars, no []:*?/\\).'''
    for ch in '[]:*?/\\':
        text = text.replace(ch, '-')
    return text[:31]


def build_combined(saved, target):
    '''Copy the sheets of each saved report into one workbook, one tab per report.

    saved: list of (tab name, file path), in tab order.
    Uses its own hidden Excel instance and always quits it.'''
    log.info(f'Combining {len(saved)} reports')
    xl = win32com.client.DispatchEx('Excel.Application')
    try:
        _fill_combined(xl, saved, target)
    finally:
        closer = excel_closer(xl)
        del xl
        closer.start()
    check_saved(target)


def _fill_combined(xl, saved, target):
    '''Build and save the combined workbook in `xl`.

    A separate function so that all its workbook and sheet references are
    released when it returns, before Excel is asked to quit.'''
    xl.Visible = False
    xl.DisplayAlerts = False
    combo = excel_call(lambda: xl.Workbooks.Add())
    blanks = [combo.Worksheets(i).Name for i in range(1, combo.Worksheets.Count + 1)]
    for name, path in saved:
        src = excel_call(lambda: xl.Workbooks.Open(os.path.abspath(path), ReadOnly=True))
        keep = [i for i in range(1, src.Worksheets.Count + 1)
                if not is_tips_sheet(src.Worksheets(i).Name)]
        for n, i in enumerate(keep, 1):
            last = combo.Worksheets(combo.Worksheets.Count)
            excel_call(lambda: src.Worksheets(i).Copy(Before=None, After=last))
            copied = combo.Worksheets(combo.Worksheets.Count)
            copied.Name = sheet_name(name if len(keep) == 1 else f'{name} {n}')
            log.debug(f'   tab "{copied.Name}" <- {os.path.basename(path)}')
        excel_call(lambda: src.Close(SaveChanges=False))
    for blank in blanks:                       # the empty sheets Add() made
        excel_call(lambda: combo.Worksheets(blank).Delete())
    excel_call(lambda: combo.Worksheets(1).Activate())
    excel_call(lambda: combo.SaveAs(os.path.abspath(target),
                                    FileFormat=EXCEL_FORMATS[FILE_EXT]))
    excel_call(lambda: combo.Close(SaveChanges=False))


# ------------------------------------------------------------- the steps ---

def qb_open_report(index, report, start, end):
    '''Step 1: open the report from Favorites and enter the dates.'''
    log.debug('Step 1: open report in QuickBooks')
    raise_window(QB_TITLE, what='QuickBooks')
    press('esc')                               # close previous report / home page
    hotkey('alt', 'r')                         # Reports menu
    press(*['down'] * FAVORITES_MENU_DOWNS, 'right')
    press(*['down'] * index, 'enter')          # pick the report
    found = wait_until(lambda: qb_report_window(report.title), QB_REPORT_TIMEOUT,
                       f'QB report window "{report.title}"', soft=True,
                       minimum=2, on_miss=log_qb_windows)
    if not found:
        raise AutoReportError(f'The "{report.title}" report window never appeared. '
                              'Wrong report opened? See the window list above.')
    time.sleep(QB_REPORT_SETTLE)               # let it finish drawing and take focus
    press('tab')                               # to first date field

    if report.date_fields == 2:
        pyautogui.write(fmt_date(start))
        log.debug(f'   typed {fmt_date(start)}')
        press('tab')
    pyautogui.write(fmt_date(end))
    log.debug(f'   typed {fmt_date(end)}')
    press('tab', *['tab'] * report.extra_tabs)  # leave field so QB refreshes
    time.sleep(2)                              # QB ignores keys while refreshing


def qb_export():
    '''Step 2: ask QuickBooks to send the report to a new Excel workbook.'''
    log.debug('Step 2: export to Excel')
    # Alt+X opens the report toolbar's Excel dropdown menu; "n" picks
    # "Create New Worksheet", which opens the "Send Report to Excel" dialog;
    # Enter exports. Never press "n" unless the menu has visibly opened: if QB
    # was busy and ignored Alt+X, the "n" lands in the report itself (it once
    # changed "Sort By" to "Num").
    before = qb_dialogs()
    for attempt in (1, 2):
        hotkey('alt', 'x')
        menu = wait_until(lambda: new_qb_windows(before, menus=True), 8,
                          f'Excel menu (attempt {attempt})', soft=True)
        if menu:
            break
        time.sleep(2)                          # QB may still be busy; try once more
                                               # (not Esc: that would close the report)
    else:
        raise AutoReportError('QuickBooks did not open the Excel menu after Alt+X '
                              '(twice). Stopped so no keys land in the report.')
    before = qb_dialogs()
    press('n')                                 # "Create New Worksheet"
    dialog = wait_until(lambda: new_qb_windows(before, menus=False), 1,
                        '"Send Report to Excel" dialog', soft=True, minimum=2)
    if dialog:   # diagnostics: confirms what the dialog looks like
        log.debug('   dialog: ' + ' | '.join(dialog.values()))
    press('enter')                             # Export


def excel_save(target, report, start, end):
    '''Step 3: check, tidy, save and close the workbook through Excel itself.'''
    log.debug('Step 3: save in Excel')
    xl, wb = wait_for_export_workbook(EXPORT_TIMEOUT)
    check_report_dates(wb, report, start, end)
    save_and_close(xl, wb, target)
    closer = excel_closer(xl)
    del xl, wb            # drop our links while Excel is still running: instant
    closer.start()        # Excel quits in the background; the run carries on
    check_saved(target)


def run_report(index, report, start, end, work_dir):
    '''Produce one report's workbook in work_dir; returns its path.'''
    log.info(f'[{index + 1}/{len(REPORTS)}] {report.title}')
    target = os.path.join(work_dir, f'{FILE_PREFIX} {file_date(end)} {report.tab}.{FILE_EXT}')
    qb_open_report(index, report, start, end)
    qb_export()
    excel_save(target, report, start, end)
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
    saved = [(report.tab, run_report(i, report, start, end, work_dir))
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
    # Working files go in the local temp folder, not OneDrive: saving there is
    # faster and OneDrive doesn't sync files that are about to be deleted.
    work_dir = tempfile.mkdtemp(prefix=f'autoreport {year}Q{quarter} {stamp} ')
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
    wait_for_excel_closers()   # let Excel finish shutting down before we exit
    if code:
        input('Press Enter to close...')
    sys.exit(code)
