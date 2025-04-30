#!/usr/bin/env python3

'''
autoreport.py
Author: James E. Pace <james@pacehouse.com>

As part of my job, I must run quarterly reports in Quickbooks, export them to Excel,
and send the Excel files to corporate.  This script attempts to automate that process.

To use, add the appropriate 5 reports to the QuickBooks favorite reports menu. The script
literally steps down that list one at a time.

Output files are saved in the hardcoded path:
~/OneDrive/Documents/Pembrook/Finance/Pembrook/auto

NB: Lots of this script is hardcoded to my specific environment and needs. It could 
be made more general, but at this point that's more work than its worth. So, caveat 
emptor.

Installation instructions: 
* Windows 10 or 11
* QuickBooks Desktop Pro 2020
* Office 365 Excel
* Install python via MS Store.  
* Install pyautogui via "pip install pyautogui"
* Install wx via "pip install wxPython"
* Install win32gui via "pip install pywin32"

Bugs / Missing Features:
* Exactly 5 reports are done, no more no less.  This should be generalized.
* Increase fault tolerance.
* Output directory is hardcoded; add it to input window instead.
* If Quickbooks or Excel move things, this will break horribly and unpredictably.
* Code should probably be modularized.
* Since reports are quarterly, autopopulate end date based on start date.

'''

import pyautogui, time, os
import wx, sys, re
import win32gui
import os   # to get the environment variables

# Debugging module
DEBUG = True
def dprint(*args, **kwargs):
    if DEBUG:
        print(*args, **kwargs)

# NB: 0.1 is too fast. 1 works. 0.3 works.
pyautogui.PAUSE = 0.3

# FIXME: Can I pull this list from QB directly?
reports = [ 'P-L YTD Comp', 'Balance Sheet Detail', 'Balance Sheet Summary', 'AP Aging Detail', 'Transaction Detail by acct' ]

# RaiseWindow(title, timeout)
# - Looks for a window of a given title and brings it to the front
# XXX: What if there are multiple windows with that title?
def RaiseWindow(title, timeout=30):
    dprint(f'-> RaiseWindow("{title}", {timeout})')
    start = time.time()
    while True:
        windows = pyautogui.getWindowsWithTitle(title)
        if windows:
            window = windows[0]
            break
        if time.time() - start > timeout:
            print(f'** ERROR: "{title}" window did not appear')
            sys.exit(1)
        time.sleep(0.2)
    window.activate()   # Bring the window to the front
    window.restore()
    win32gui.SetForegroundWindow(window._hWnd)
    dprint(f'<- RaiseWindow("{title}", {timeout})')

# WaitCloseWindow (title, timeout)
# - Waits until a window with name 'title' does not exist (or times out)
# - Returns true or false, as appropriate
def WaitCloseWindow (title, timeout=30):
    dprint(f'-> WaitCloseWindow("{title}", {timeout})')
    start = time.time()
    while True:
        windows = pyautogui.getWindowsWithTitle(title)
        if not windows:
            dprint(f'<- WaitCloseWindow("{title}", {timeout}) - TRUE')
            return True
        if time.time() - start > timeout:
            print(f'** ERROR: "{title}" window did not close')
            dprint(f'<- WaitCloseWindow("{title}", {timeout}) - FALSE')
            return False
        time.sleep(0.2)

# GetInput: Dialog box that asks the user to put the range of dates to use for reports.
class GetInput(wx.Dialog):
    def __init__(self, parent):
        # XXX: 4/28/25: Window size broken on new higher rez monitor, so redo the size and positions here
        wx.Dialog.__init__(self, parent, wx.ID_ANY, "QuickBooks Report Range", size=(500, 320))
        self.panel = wx.Panel(self, wx.ID_ANY)

        self.lblname = wx.StaticText(self.panel, label="Start Date", pos=(40, 30))
        self.StartDate = wx.TextCtrl(self.panel, value="", pos=(150, 30), size=(250, -1))
        self.lblsur = wx.StaticText(self.panel, label="End Date", pos=(40, 90))
        self.EndDate = wx.TextCtrl(self.panel, value="", pos=(150, 90), size=(250, -1))
        self.StartButton = wx.Button(self.panel, label="Start", pos=(110, 160), size=(120, 50))
        self.QuitButton = wx.Button(self.panel, label="Quit", pos=(270, 160), size=(120, 50))

        self.StartButton.Bind(wx.EVT_BUTTON, self.OnStart)
        self.QuitButton.Bind(wx.EVT_BUTTON, self.OnQuit)
        self.Bind(wx.EVT_CLOSE, self.OnQuit)
        self.SetDefaultItem(self.StartButton)

        self.Centre()
        self.SetMinSize((500, 320))
        self.Show()
        self.StartDate.SetFocus()

    def OnQuit(self, event):
        self.Destroy()
        sys.exit()

    def OnStart(self, event):
        global startDate, endDate, fileDate
        self.StartDate = self.StartDate.GetValue()
        self.EndDate = self.EndDate.GetValue()
        # Validate the dates are formatted properly, get filename date, and save them globally
        regex = re.compile( r"^\d+/\d+/\d+$")
        if regex.match (self.StartDate):
            startDate = self.StartDate
            dprint ("Start Date: " + startDate + "\n")
        else:
            # FIXME: Loop until valid date
            startDate = ""
            print("Error: Bad start date format: " + self.StartDate)
            
        if regex.match (self.EndDate):
            endDate = self.EndDate
            sp = endDate.split("/")
            fileDate = str(int(sp[2])%100).zfill(2) + sp[0].zfill(2) + sp[1].zfill(2)
            dprint ("End Date: " + endDate + " [" + fileDate + "]\n")
        else:
            # FIXME: Loop until valid date. Does this just go now?
            print("Error: Bad end date format: " + self.EndDate)
            time.sleep(10)
            endDate = ""
            fileDate = ""
        self.Destroy()

# Get the date input
app = wx.App()
dlg = GetInput(None) 
dlg.ShowModal()

print(f'Start Date: {startDate}')
print(f'End Date: {endDate}')
print(f'File Date: {fileDate}')

if not startDate or not endDate:
    print(f'ERROR: User date input not valid. Exiting.')
    time.sleep(10)
    sys.exit()

# Get OneDrive path from environment
# XXX: 4/28/25: Moved directory to OneDrive
# FIXME: Don't hardcode the path. Choose in the input window?
onedrive = os.environ["OneDrive"]
outputDir = os.path.join(onedrive, "Documents", "Pembrook", "Finance", "Pembrook", "auto")
print(f'Output directory: {outputDir}')

# Make sure output directory exists
if not os.path.exists(outputDir):
    os.makedirs (outputDir)

# Begin looping over the reports
# FIXME: range(0,5) should be replaced with iterator of reports array
for i in range (0, 5):
    # Let's attempt to bring focus to Quickbooks
    RaiseWindow("QuickBooks Desktop Pro 2020")
    dprint("QB Foregrounded")

    print (f'Loop {str(i)}: {reports[i]}')
    time.sleep(0.5)

    # Run the report
    pyautogui.press('esc')          # Clear previous report, if there (or close home page)
    pyautogui.hotkey('alt', 'r')    # Open Reports menu
    pyautogui.press('down')         # To memorized reports
    pyautogui.press('down')         # To scheduled reports
    pyautogui.press('down')         # To commented reports
    pyautogui.press('down')         # To Favorite Reports (yay!)
    pyautogui.press('right')        # Enter Favorite Reports
    for j in range (0, i):
        pyautogui.press ('down')    # Move to desired report
    pyautogui.press('enter')        # Select desired report
    time.sleep(2)                   # Give time for desired report to load
    pyautogui.press('tab')          # Move to the Starting Date field

    # NB: "Balance Sheet Summary" [2] and "A/P Aging" [3] only have 1 date field
    # FIXME: Don't hard code which are special
    if (i != 2 and i != 3 ):
        pyautogui.typewrite(startDate)
        pyautogui.press('tab')      # Move to Ending Date field

    pyautogui.typewrite(endDate)
    pyautogui.press('tab')          # Move out of date field so it registers

    # XXX: Fuck Intuit.  To update the A/P Aging report, you've got to tab a 
    # couple more times.
    if (i == 3):
        pyautogui.press('tab')
        pyautogui.press('tab')

    time.sleep(1)                   # Give a chance to catch up

    # Ask for Excel version of report
    pyautogui.hotkey('alt', 'x')
    pyautogui.press('n')
    time.sleep (2)
    pyautogui.press('enter')

    # No extension here as it is added automatically 
    filename = os.path.join(outputDir, "Pembrook " + fileDate + " " + reports[i])
    
    # Delete file before trying to create it
    # NB: Since we don't have the extension, let's nuke a couple options
    for ext in ("xlsx", "xlsm"):
        if os.path.isfile(filename+"."+ext):
            os.remove(filename+"."+ext)
            dprint(f'Removed: "{filename}.{ext}"')
    print (f'Report: "{filename}"')

    RaiseWindow("Excel", 120)    # Give long delay for Excel, as it might take a while
    
    # In Excel, Save via the Save As window
    time.sleep(0.5)
    pyautogui.press('f12')                  # Ask for the Save As window
    #time.sleep(0.5)
    RaiseWindow("Save As")                  # Switch to to the Save As window
    time.sleep(0.5)
    pyautogui.typewrite(f"{filename}")      # Write the filename in 'File name'
    dprint(f'Typewrote "{filename}"')
    time.sleep(0.5)
    pyautogui.press('tab')                  # Move to 'Save as type'
    dprint("tab")
    pyautogui.press('down')                 # Activate type dropdown
    dprint("down")
    pyautogui.press('home')                 # Go to the top of the list, xlsx
    dprint("home")
    pyautogui.press('down')                 # Move from xlsx to xlsm
    dprint("down")
    pyautogui.press('enter')                # Select xlsm
    dprint("enter")
    pyautogui.hotkey('alt','s')             # Save the file
    dprint("alt-s")
    
    if WaitCloseWindow("Save As", 60):
        dprint(f'Saved "{filename}"')
    else:
        print(f'** ERROR: Window "Save As" did not close')
        sys.exit(1)
    time.sleep(1)
    
    # Exit Excel
    RaiseWindow("Excel")                    # Make sure Excel is our focus
    time.sleep(15)                          # Sometimes, it is still processing, which breaks things. OneDrive?
    pyautogui.hotkey('alt', 'f4')           # Exit Excel
    dprint("alt-f4")
    time.sleep(1)

    if WaitCloseWindow("Excel", 60):
        dprint(f'Done with "{filename}"')
    else:
        print(f'** ERROR: Window "Excel" did not close')
        sys.exit(1)
    dprint("Excel closed")
# End Loop
dprint ("Done with all reports!")

# Reopen the home screen in Quickbooks
RaiseWindow("QuickBooks Desktop Pro 2020")
pyautogui.press('esc')          # Clear previous report
pyautogui.hotkey('alt', 'c')    # Open the Company menu
pyautogui.press('enter')        # Open the Home Page

print ("REPORTS GENERATED SUCCESSFULLY")