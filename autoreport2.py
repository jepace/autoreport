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
* Dates are validated as real calendar dates; the dialog stays open and
  shows an error until the input is valid.  End date auto-fills to the end
  of the quarter that starts on the start date.
* The output folder can be chosen in the dialog.
* Failures raise exceptions with a clear message; the console window waits
  for Enter instead of vanishing.
* Instead of a fixed 10 second sleep, waits for the saved file to appear on
  disk and stop growing before closing Excel.
* Handles a "Save changes?" prompt when closing Excel.
* File paths are pasted via the clipboard (safe for non-ASCII characters).

Requirements: Windows, QuickBooks Desktop, Excel, and
    pip install pyautogui wxPython pywin32

Emergency stop: slam the mouse into a screen corner (pyautogui FAILSAFE).
'''

import datetime as dt
import os
import sys
import time
from dataclasses import dataclass

import pyautogui
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

QB_TITLE = 'QuickBooks Desktop Pro 2020'
EXCEL_TITLE = 'Excel'
SAVE_AS_TITLE = 'Save As'

# Number of "down" presses in the Reports menu to reach "Favorite Reports"
FAVORITES_MENU_DOWNS = 4

FILE_PREFIX = 'Pembrook'
FILE_EXT = 'xlsm'          # The script picks the 2nd entry in "Save as type"

def default_output_dir():
    base = os.environ.get('OneDrive') or os.path.expanduser('~')
    return os.path.join(base, 'Documents', 'Pembrook', 'Finance', 'Pembrook', 'auto')

DATE_FMT = '%m/%d/%Y'      # Format typed into QuickBooks

pyautogui.PAUSE = 0.3      # 0.1 is too fast for QB; 0.3 works
DEBUG = True

# ============================================================================


class AutoReportError(Exception):
    pass


def dprint(*args, **kwargs):
    if DEBUG:
        print(*args, **kwargs)


# ---------------------------------------------------------------- dates ----

def parse_date(text):
    '''Parse m/d/yyyy (or m/d/yy). Returns a date or None.'''
    text = text.strip()
    for fmt in ('%m/%d/%Y', '%m/%d/%y'):
        try:
            return dt.datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return None


def end_of_quarter(start):
    '''Last day of the 3-month period beginning on start.'''
    month = start.month + 3
    year = start.year + (month - 1) // 12
    month = (month - 1) % 12 + 1
    return dt.date(year, month, 1) - dt.timedelta(days=1)


def fmt_date(d):
    # QB accepts unpadded m/d/yyyy, matching what users typed in v1
    return f'{d.month}/{d.day}/{d.year}'


def file_date(d):
    '''YYMMDD, matching the original script's file naming.'''
    return d.strftime('%y%m%d')


# ---------------------------------------------------------------- dialog ---

class InputDialog(wx.Dialog):
    def __init__(self):
        super().__init__(None, title='QuickBooks Report Range',
                         style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        self.start = self.end = self.output_dir = None
        self._end_autofilled = True

        self.start_ctrl = wx.TextCtrl(self)
        self.end_ctrl = wx.TextCtrl(self)
        self.dir_ctrl = wx.DirPickerCtrl(self, path=default_output_dir(),
                                         style=wx.DIRP_USE_TEXTCTRL)
        hint = wx.StaticText(self, label='Dates as m/d/yyyy. End date fills in '
                                         'automatically for a quarter.')

        grid = wx.FlexGridSizer(cols=2, vgap=10, hgap=10)
        grid.AddGrowableCol(1)
        for label, ctrl in (('Start Date', self.start_ctrl),
                            ('End Date', self.end_ctrl),
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
        main.Add(hint, 0, wx.LEFT | wx.RIGHT, 15)
        main.Add(buttons, 0, wx.EXPAND | wx.ALL, 15)
        self.SetSizerAndFit(main)
        self.SetSize((max(self.GetSize()[0], 520), -1))

        start_btn.SetDefault()
        start_btn.Bind(wx.EVT_BUTTON, self.on_start)
        self.start_ctrl.Bind(wx.EVT_TEXT, self.on_start_changed)
        self.end_ctrl.Bind(wx.EVT_TEXT, self.on_end_changed)
        self.Centre()
        self.start_ctrl.SetFocus()

    def on_start_changed(self, event):
        start = parse_date(self.start_ctrl.GetValue())
        if start and self._end_autofilled:
            self.end_ctrl.ChangeValue(fmt_date(end_of_quarter(start)))  # no EVT_TEXT
        event.Skip()

    def on_end_changed(self, event):
        self._end_autofilled = not self.end_ctrl.GetValue()  # user took over
        event.Skip()

    def _error(self, msg, ctrl):
        wx.MessageBox(msg, 'Invalid input', wx.OK | wx.ICON_ERROR, self)
        ctrl.SetFocus()
        ctrl.SelectAll()

    def on_start(self, event):
        start = parse_date(self.start_ctrl.GetValue())
        end = parse_date(self.end_ctrl.GetValue())
        out = self.dir_ctrl.GetPath().strip()
        if not start:
            return self._error('Start date is not a valid date (m/d/yyyy).', self.start_ctrl)
        if not end:
            return self._error('End date is not a valid date (m/d/yyyy).', self.end_ctrl)
        if end < start:
            return self._error('End date is before start date.', self.end_ctrl)
        if not out:
            return wx.MessageBox('Choose an output folder.', 'Invalid input',
                                 wx.OK | wx.ICON_ERROR, self)
        self.start, self.end, self.output_dir = start, end, out
        self.EndModal(wx.ID_OK)


def get_input():
    '''Returns (start, end, output_dir) or None if the user quit.'''
    app = wx.App(False)  # noqa: F841  (must stay alive while dialog runs)
    dlg = InputDialog()
    try:
        if dlg.ShowModal() != wx.ID_OK:
            return None
        return dlg.start, dlg.end, dlg.output_dir
    finally:
        dlg.Destroy()


# --------------------------------------------------------------- windows ---

def find_window(title):
    windows = pyautogui.getWindowsWithTitle(title)
    if len(windows) > 1:
        dprint(f'   note: {len(windows)} windows match "{title}", using first')
    return windows[0] if windows else None


def raise_window(title, timeout=30):
    dprint(f'-> raise_window("{title}")')
    deadline = time.time() + timeout
    while (window := find_window(title)) is None:
        if time.time() > deadline:
            raise AutoReportError(f'"{title}" window did not appear within {timeout}s')
        time.sleep(0.2)
    if window.isMinimized:
        window.restore()
    try:
        window.activate()
    except Exception:  # pygetwindow can raise even when activation worked
        pass
    win32gui.SetForegroundWindow(window._hWnd)
    time.sleep(0.3)


def wait_window_closed(title, timeout=30):
    deadline = time.time() + timeout
    while find_window(title) is not None:
        if time.time() > deadline:
            return False
        time.sleep(0.2)
    return True


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


def paste_text(text):
    '''Type text via the clipboard (handles any characters).'''
    app = wx.GetApp() or wx.App(False)  # noqa: F841
    if wx.TheClipboard.Open():
        try:
            wx.TheClipboard.SetData(wx.TextDataObject(text))
        finally:
            wx.TheClipboard.Close()
        pyautogui.hotkey('ctrl', 'v')
    else:
        pyautogui.write(text)  # fallback


def press(*keys):
    for k in keys:
        pyautogui.press(k)
    dprint('   keys:', ' '.join(keys))


# ------------------------------------------------------------- the work ----

def run_report(index, report, start, end, output_dir):
    print(f'[{index + 1}/{len(REPORTS)}] {report.name}')
    base = os.path.join(output_dir, f'{FILE_PREFIX} {file_date(end)} {report.name}')
    target = f'{base}.{FILE_EXT}'

    # Remove old outputs first, so a stale file can't be mistaken for success
    for ext in ('xlsx', 'xlsm'):
        old = f'{base}.{ext}'
        if os.path.isfile(old):
            try:
                os.remove(old)
                dprint(f'   removed old {old}')
            except PermissionError:
                raise AutoReportError(f'Cannot delete {old} -- is it open in Excel?')

    # --- QuickBooks: open the report and set the dates
    raise_window(QB_TITLE)
    press('esc')                               # close previous report / home page
    pyautogui.hotkey('alt', 'r')               # Reports menu
    press(*['down'] * FAVORITES_MENU_DOWNS, 'right')
    press(*['down'] * index, 'enter')          # pick the report
    time.sleep(2)
    press('tab')                               # to first date field

    if report.date_fields == 2:
        pyautogui.write(fmt_date(start))
        press('tab')
    pyautogui.write(fmt_date(end))
    press('tab', *['tab'] * report.extra_tabs)  # leave field so QB refreshes
    time.sleep(1)

    # --- Export to a new Excel workbook
    pyautogui.hotkey('alt', 'x')
    press('n')
    time.sleep(2)
    press('enter')

    # --- Excel: Save As .xlsm
    raise_window(EXCEL_TITLE, timeout=300)     # Excel can be slow to start
    press('f12')
    raise_window(SAVE_AS_TITLE)
    paste_text(base)                           # extension added by type choice
    press('tab', 'down', 'home', 'down', 'enter')  # Save as type -> 2nd entry (xlsm)
    pyautogui.hotkey('alt', 's')
    if not wait_window_closed(SAVE_AS_TITLE, 90):
        raise AutoReportError('"Save As" window did not close')

    wait_file_stable(target)
    print(f'   saved {target}')

    # --- Close Excel
    raise_window(EXCEL_TITLE)
    pyautogui.hotkey('alt', 'f4')
    if not wait_window_closed(EXCEL_TITLE, 15):
        # Probably a "Save changes?" prompt; file is already saved, so don't save.
        dprint('   Excel still open; answering "Don\'t Save"')
        pyautogui.hotkey('alt', 'n')
        if not wait_window_closed(EXCEL_TITLE, 45):
            raise AutoReportError('Excel did not close')


def main():
    user = get_input()
    if user is None:
        print('Cancelled.')
        return 0
    start, end, output_dir = user
    print(f'Range: {fmt_date(start)} - {fmt_date(end)}   Output: {output_dir}')
    os.makedirs(output_dir, exist_ok=True)

    for i, report in enumerate(REPORTS):
        run_report(i, report, start, end, output_dir)

    # Back to the QuickBooks home page
    raise_window(QB_TITLE)
    press('esc')
    pyautogui.hotkey('alt', 'c')
    press('enter')
    print('REPORTS GENERATED SUCCESSFULLY')
    return 0


if __name__ == '__main__':
    try:
        code = main()
    except AutoReportError as e:
        print(f'\n** ERROR: {e}')
        code = 1
    except pyautogui.FailSafeException:
        print('\n** Stopped by user (mouse moved to screen corner).')
        code = 1
    if code:
        input('Press Enter to close...')
    sys.exit(code)
