# Code Issues Report — autoreport.py

## Critical Bugs

### 1. Global variables used before assignment (lines 153–155)
`startDate`, `endDate`, and `fileDate` are never initialised before the dialog
runs. If the user closes the dialog via the window manager's X button — which
fires `OnQuit` → `sys.exit()` — that is fine; but if anything raises an
exception before `OnStart` stores values into them, or if the dialog is
dismissed in any other way, lines 153–155 will raise a `NameError`.

**Fix:** initialise the three globals to `""` near the top of the script,
before `app = wx.App()`.

---

### 2. Widget self-overwrite in `OnStart` (lines 123–124)
```python
self.StartDate = self.StartDate.GetValue()   # overwrites wx.TextCtrl with str
self.EndDate   = self.EndDate.GetValue()     # same
```
`self.StartDate` is first the `wx.TextCtrl` widget; after line 123 it becomes
a plain string. This works by accident because `.GetValue()` is called before
the widget reference is lost, but it is deeply confusing and fragile. Any
future access of `self.StartDate` will get a string, not a widget.

**Fix:** use distinct names, e.g. `start_val = self.StartDate.GetValue()`.

---

### 3. `Show()` + `ShowModal()` race (lines 114, 151)
`GetInput.__init__` calls `self.Show()` (non-modal), and then the caller
immediately calls `dlg.ShowModal()`. On some platforms `ShowModal()` after
`Show()` produces unpredictable behaviour; on others the script continues past
`ShowModal()` immediately (before the user has typed anything), causing a
`NameError` on `startDate`.

**Fix:** remove `self.Show()` from `__init__`; `ShowModal()` is enough.

---

### 4. `sys.exit()` inside a wx event handler (line 119)
```python
def OnQuit(self, event):
    self.Destroy()
    sys.exit()
```
`sys.exit()` raises `SystemExit` mid-event, bypassing wx's event-loop cleanup
and potentially leaving the GUI in an inconsistent state. On Windows this can
cause a crash or hung process.

**Fix:** call `wx.GetApp().ExitMainLoop()` and then return, letting wx shut
down normally.

---

### 5. `time.sleep(10)` inside the GUI event handler (line 143)
When the end-date regex fails, the code sleeps for 10 seconds inside `OnStart`
before proceeding. This freezes the entire UI (wx runs on the main thread) with
no visual feedback to the user.

**Fix:** show a `wx.MessageDialog` with the error, then return without
destroying the dialog so the user can correct the input.

---

## Correctness Issues

### 6. Duplicate `import os` (lines 39, 42)
`os` is imported twice. The second import is a no-op but is misleading.

---

### 7. Overly permissive date regex (line 126)
```python
regex = re.compile(r"^\d+/\d+/\d+$")
```
This accepts `99/99/9999`, `0/0/0`, etc. Invalid dates pass validation and are
sent to QuickBooks, where they will silently produce wrong reports.

**Fix:** parse with `datetime.datetime.strptime(value, "%m/%d/%Y")` inside a
`try/except ValueError`.

---

### 8. Missing `OneDrive` environment variable raises `KeyError` (line 165)
```python
onedrive = os.environ["OneDrive"]
```
On machines where OneDrive is not installed or the variable has a different
name, this raises an unhandled `KeyError` with no user-friendly message.

**Fix:** use `os.environ.get("OneDrive")` and handle the `None` case
explicitly.

---

### 9. 2-digit year arithmetic is fragile (line 138)
```python
str(int(sp[2]) % 100).zfill(2)
```
For the year 2100 this produces `"00"`, colliding with year-2000 filenames.
The `%100` truncation is not documented and silently changes behaviour when
the year rolls over a century boundary.

---

### 10. `pyautogui.typewrite()` silently drops non-ASCII characters (lines 210, 215, 259)
`typewrite` only supports ASCII. If the file path (`outputDir`, report name, or
date string) ever contains a non-ASCII character, those characters are silently
skipped, corrupting the typed input without any error.

**Fix:** use `pyautogui.write()` with the `interval` parameter, or use the
clipboard approach (`pyperclip` + Ctrl+V) for robustness.

---

### 11. Stale file deletion does not account for in-use files (lines 244–247)
```python
for ext in ("xlsx", "xlsm"):
    if os.path.isfile(filename+"."+ext):
        os.remove(filename+"."+ext)
```
If Excel has the file open (from a previous failed run), `os.remove()` will
raise a `PermissionError` on Windows, crashing the script with no recovery
path.

---

## Design / Maintainability Issues

### 12. No `if __name__ == "__main__":` guard (entire script)
All top-level logic executes on import. This prevents unit testing and makes
the script impossible to import without side effects.

---

### 13. Hardcoded company name and path (line 166)
```python
outputDir = os.path.join(onedrive, "Documents", "Pembrook", "Finance", "Pembrook", "auto")
```
Acknowledged as a FIXME. The output directory should be configurable via the
`GetInput` dialog or a config file.

---

### 14. Hardcoded report count and menu-navigation logic (lines 175, 198–200)
The script assumes exactly 5 reports and navigates the QuickBooks menu by
pressing `down` `i` times. If QuickBooks changes the menu layout, adds a
separator, or if the favourites list order changes, the wrong report will be
silently selected and exported.

---

### 15. Hardcoded special-case report names (line 209)
```python
if report not in ("Balance Sheet Summary", "AP Aging Detail"):
```
These names must match the QB favourites list exactly (case-sensitive). A
typo in the `reports` list would silently cause the wrong number of date
fields to be filled.

---

### 16. 10-second sleep before closing Excel (line 284)
```python
time.sleep(10)   # Sometimes, it is still processing, which breaks things. OneDrive?
```
This unconditional wait slows every report by 10 seconds. The root cause
(OneDrive sync locking the file) is not addressed; the sleep may or may not be
long enough.

---

### 17. `WaitCloseWindow("Excel")` does not handle "do you want to save?" prompts (line 289)
After `alt+f4`, Excel may show a save-confirmation dialog. The title of that
dialog is not "Excel", so `WaitCloseWindow` will time out after 60 seconds and
exit with an error rather than dismissing the dialog.

---

### 18. Typo in comment (line 252)
```
# Switch to to the Save As window
```
"to to" should be "to".

---

## Summary Table

| # | Severity | Category | Location |
|---|----------|----------|----------|
| 1 | Critical | Bug | lines 153–155 |
| 2 | Critical | Bug | lines 123–124 |
| 3 | Critical | Bug | lines 114, 151 |
| 4 | High | Reliability | line 119 |
| 5 | High | UX / Bug | line 143 |
| 6 | Low | Code quality | lines 39, 42 |
| 7 | High | Correctness | line 126 |
| 8 | High | Reliability | line 165 |
| 9 | Medium | Correctness | line 138 |
| 10 | Medium | Correctness | lines 210, 215, 259 |
| 11 | Medium | Reliability | lines 244–247 |
| 12 | Low | Maintainability | entire script |
| 13 | Low | Maintainability | line 166 |
| 14 | Medium | Reliability | lines 175, 198–200 |
| 15 | Medium | Reliability | line 209 |
| 16 | Low | Performance | line 284 |
| 17 | Medium | Reliability | line 289 |
| 18 | Low | Documentation | line 252 |
