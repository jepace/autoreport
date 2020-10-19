import pyautogui, time, os
import wx, sys, re

# NB: 0.1 is too fast. 1 works. Does 0.25?
pyautogui.PAUSE = 0.25

reports = [ 'P-L YTD Comp', 'Balance Sheet Detail', 'Balance Sheet Summary', 'AP Aging Detail', 'Transaction Detail by acct' ]

# FIXME: Ask for dates. What's the scope of these variables?
global startDate, endDate, fileDate

class MyFrame(wx.Frame):
    def __init__(self, parent):
        wx.Frame.__init__(self, parent, -1, "Output",size=(500,400))
        self.panel = wx.Panel(self)
        sizer = wx.BoxSizer(wx.VERTICAL)
        self.log = wx.TextCtrl(self.panel, wx.ID_ANY, size=(400,300),style = wx.TE_MULTILINE|wx.TE_READONLY|wx.VSCROLL)
        self.button = wx.Button(self.panel, label="Done")
        sizer.Add(self.log, 0, wx.EXPAND | wx.ALL, 10)
        sizer.Add(self.button, 0, wx.EXPAND | wx.ALL, 10)
        self.panel.SetSizer(sizer)
        self.Bind(wx.EVT_BUTTON, self.OnButton)
        
    def OnButton(self,event):
        self.Destroy()
        sys.exit()

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
            # frame.log.AppendText ("Start Date: " + startDate + "\n")
        else:
            startDate = ""
            frame.log.AppendText("Error: Bad start date format: " + self.StartDate + "\n" )
        if regex.match (self.EndDate):
            endDate = self.EndDate
            sp = endDate.split("/")
            fileDate = str(int(sp[2])%100).zfill(2) + sp[0].zfill(2) + sp[1].zfill(2)
            # frame.log.AppendText ("End Date: " + endDate + " [" + fileDate + "]\n")
        else:
            frame.log.AppendText("Error: Bad end date format: " + self.EndDate + "\n" )
            endDate = ""
            fileDate = ""
        self.Destroy()

# Create the output window, named 'frame'
app = wx.App()
frame = MyFrame(None)
frame.Show()

# Get the date input
dlg = GetInput(parent = frame.panel) 
dlg.ShowModal()
frame.log.AppendText("[M] Start Date: " + startDate + "\n")
frame.log.AppendText("[M] End Date: " + endDate + "\n")
frame.log.AppendText("[M] File Date: " + fileDate + "\n")

if not startDate or not endDate:
    frame.log.AppendText("User input not found. Exiting.\n")
    sys.exit()

# app.MainLoop()
"""
Pay special attention to the result items
"""

# Make sure output directory exists
outputDir = 'C:\\Users\\jepace\\Documents\\Pembrook\\Finance\\Pembrook\\auto\\'
if not os.path.exists(outputDir):
    os.makedirs (outputDir)

frame.log.AppendText("Ensure QuickBooks is open. Be sure all instances of Excel are closed.\n")
frame.log.AppendText('>>> 5 SECOND PAUSE.  Switch to Quickbooks or hit Ctrl-C <<<\n')
time.sleep(5)

# FIXME: Select Quickbooks (and Excel) for real, not just assume they will be in focus
#hwnd = win32gui.FindWindow("","QuickBooks")
#win32gui.SetFocus(hwmd)

for i in range (0, 5):
    frame.log.AppendText ("Loop: " + reports[i] + "\n")
    frame.Update()

    # Run the report
    pyautogui.hotkey('alt', 'r')
    pyautogui.hotkey('down')
    pyautogui.hotkey('down')
    pyautogui.hotkey('down')
    pyautogui.hotkey('right')
    for j in range (0, i):
        pyautogui.hotkey ('down')
    pyautogui.hotkey('enter')
    pyautogui.hotkey('tab')

    # NB: Balance Sheet Summary and A/P Aging only have 1 date field
    if (i != 2 and i != 3 ):
        pyautogui.typewrite(startDate)
        pyautogui.hotkey('tab')

    pyautogui.typewrite(endDate)
    pyautogui.hotkey('tab')

    # Ask for Excel version of report
    pyautogui.hotkey('alt', 'x')
    pyautogui.hotkey('n')
    time.sleep (2)
    pyautogui.hotkey('enter')

    # Delete file before trying to create it
    filename = outputDir+"Pembrook "+fileDate+" "+reports[i]+".xlsx"
    if os.path.isfile(filename):
        os.remove(filename)
    frame.log.AppendText ("Report: " + filename + "\n")

    # Wait for Excel to launch
    frame.log.AppendText ("Sleeping while Excel loads...")
    frame.Update()
    time.sleep (90)
    frame.log.AppendText ("Awake!\n")
    
    # In Excel, save the report
    pyautogui.hotkey('F12')
    time.sleep (1)
    pyautogui.typewrite(filename)
    pyautogui.hotkey('enter')
    pyautogui.hotkey('alt', 'F4')

    # Should be back in Quickbooks.  Clear the report away and do it again.
    time.sleep (1)
    pyautogui.hotkey('esc')
# End Loop
frame.log.AppendText ("Done!\n")
