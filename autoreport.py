import pyautogui, time, os
import wx, sys, re
import win32gui
import os   # to get the environment variables

# Installation instructions: 
# On Windows 10, install python via MS Store.  
# Install pyautogui via "pip install pyautogui"
# Install wx via "pip install wxPython"
# Install win32gui via "pip install pywin32"

# NB: 0.1 is too fast. 1 works. Does 0.25?
pyautogui.PAUSE = 0.3

reports = [ 'P-L YTD Comp', 'Balance Sheet Detail', 'Balance Sheet Summary', 'AP Aging Detail', 'Transaction Detail by acct' ]

class GetInput(wx.Dialog):
    def __init__(self, parent):
        wx.Dialog.__init__(self, parent, wx.ID_ANY, "QuickBooks Report Range", size= (250,180))
        self.panel = wx.Panel(self,wx.ID_ANY)

        self.lblname = wx.StaticText(self.panel, label="Start Date", pos=(20,20))
        self.StartDate = wx.TextCtrl(self.panel, value="", pos=(90,20), size=(110,-1))
        self.lblsur = wx.StaticText(self.panel, label="End Date", pos=(20,60))
        self.EndDate = wx.TextCtrl(self.panel, value="", pos=(90,60), size=(110,-1))
        self.StartButton =wx.Button(self.panel, label="Start", pos=(20,100))
        self.QuitButton =wx.Button(self.panel, label="Quit", pos=(110,100))
        self.StartButton.Bind(wx.EVT_BUTTON, self.OnStart)
        self.QuitButton.Bind(wx.EVT_BUTTON, self.OnQuit)
        self.Bind(wx.EVT_CLOSE, self.OnQuit)
        self.Show()

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
            # print ("Start Date: " + startDate + "\n")
        else:
            startDate = ""
            print("Error: Bad start date format: " + self.StartDate)

        if regex.match (self.EndDate):
            endDate = self.EndDate
            sp = endDate.split("/")
            fileDate = str(int(sp[2])%100).zfill(2) + sp[0].zfill(2) + sp[1].zfill(2)
            # print ("End Date: " + endDate + " [" + fileDate + "]\n")
        else:
            print("Error: Bad end date format: " + self.EndDate)
            time.sleep(10)
            endDate = ""
            fileDate = ""
        self.Destroy()

# Get the date input
app = wx.App()
dlg = GetInput(None) 
dlg.ShowModal()

print("[M] Start Date: " + startDate)
print("[M] End Date: " + endDate)
print("[M] File Date: " + fileDate)

if not startDate or not endDate:
    print("** ERROR ** User input not valid. Exiting.")
    time.sleep(10)
    sys.exit()

# Make sure output directory exists
drive = os.environ["HOMEDRIVE"]
homedir = os.environ["HOMEPATH"]
outputDir = drive + homedir + '\\Documents\\Pembrook\\Finance\\Pembrook\\auto\\'
print ("Output directory: " + outputDir)

if not os.path.exists(outputDir):
    os.makedirs (outputDir)

#print("Ensure QuickBooks is open. Be sure all instances of Excel are closed.")
#print('>>> 5 SECOND PAUSE: Switch to Quickbooks or hit Ctrl-C <<<')
#time.sleep(5)

for i in range (0, 5):
    # Let's attempt to bring focus to Quickbooks
    hwnd = pyautogui.getWindowsWithTitle("QuickBooks Desktop Pro 2020")[0]
    hwnd.restore()
    win32gui.SetForegroundWindow(hwnd._hWnd)

    print ("Loop " + str(i) + ": " + reports[i])
    time.sleep(1)

    pyautogui.hotkey('esc')     # Clear previous report, if there

    # Run the report
    pyautogui.hotkey('alt', 'r')
    pyautogui.hotkey('down')
    pyautogui.hotkey('down')
    pyautogui.hotkey('down')
    pyautogui.hotkey('down')    # QB2020 added another menu option
    pyautogui.hotkey('right')
    for j in range (0, i):
        pyautogui.hotkey ('down')
    pyautogui.hotkey('enter')
    time.sleep(2)
    pyautogui.hotkey('tab')

    # NB: "Balance Sheet Summary" [2] and "A/P Aging" [3] only have 1 date field
    if (i != 2 and i != 3 ):
        pyautogui.typewrite(startDate)
        pyautogui.hotkey('tab')

    pyautogui.typewrite(endDate)
    pyautogui.hotkey('tab')

    # Fuck Intuit.  To update the A/P Aging report, you've got to tab a 
    # couple more times.
    if (i == 3):
        pyautogui.hotkey('tab')
        pyautogui.hotkey('tab')

    time.sleep(1)

    # Ask for Excel version of report
    pyautogui.hotkey('alt', 'x')
    pyautogui.hotkey('n')
    time.sleep (2)
    pyautogui.hotkey('enter')

    # Delete file before trying to create it
    # 7/29/22: Removing extension? 
    # filename = outputDir+"Pembrook "+fileDate+" "+reports[i]+".xlsm"
    filename = outputDir+"Pembrook "+fileDate+" "+reports[i]
    if os.path.isfile(filename):
        os.remove(filename)
    print ("Report: " + filename)

    # Wait for Excel to launch.  For the large report, it really does
    # take a long time, so 90 secs is right.  (This makes debugging a
    # pain, so you might want to alter it for testing.)
    print ("Sleeping a long time while Excel loads...")
    time.sleep (20)
    #print ("Awake!")
    
    # Switch to the Excel program.
    hwnd = pyautogui.getWindowsWithTitle("Excel")[0]
    win32gui.SetForegroundWindow(hwnd._hWnd)

    # In Excel, save the report
    # XXX 7/25/22: F12 leads to macro problems during save
    pyautogui.hotkey('F12')
    time.sleep (1)
    pyautogui.typewrite(filename)
    pyautogui.hotkey('enter')
    pyautogui.hotkey('enter')   # To get past some dialog about macros in 2020
    
    # Exit Excel (FIXME 201019 - Excel not closing; increase this
    # sleep from 2 to 5 secs)
    time.sleep (5)
    pyautogui.hotkey('alt', 'F4')
# End Loop

#print ("Done!")

# Reopen the home screen in Quickbooks
hwnd = pyautogui.getWindowsWithTitle("QuickBooks Desktop Pro 2020")[0]
win32gui.SetForegroundWindow(hwnd._hWnd)
pyautogui.hotkey('esc')     # Clear previous report, if there
pyautogui.hotkey('alt', 'c')
pyautogui.hotkey('enter')

print ("And we're out of here!")
time.sleep (3)
