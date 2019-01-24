import pyautogui, time, os

# NB: 0.1 is too fast. 1 works. Does 0.25?
pyautogui.PAUSE = 0.25

reports = [ 'P-L YTD Comp', 'Balance Sheet Detail', 'Balance Sheet Summary', 'AP Aging Detail', 'Transaction Detail by acct' ]

# FIXME: Ask for dates
startDate = "09/01/2018"
endDate = "12/31/2018"
endDateClean = "181231"

# Make sure output directory exists
outputDir = 'C:\\Users\\jepace\\Documents\\Pembrook\\Finance\\Pembrook\\auto\\'
if not os.path.exists(outputDir):
    os.makedirs (outputDir)

print('>>> 5 SECOND PAUSE.  Switch to Quickbooks or hit Ctrl-C <<<')
time.sleep(5)

# FIXME: Select Quickbooks (and Excel) for real, not just assume they will be in focus
#hwnd = win32gui.FindWindow("","QuickBooks")
#win32gui.SetFocus(hwmd)

for i in range (0, 5):
    print ("Loop: ", reports[i])

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
    filename = outputDir+"Pembrook "+endDateClean+" "+reports[i]+".xlsx"
    if os.path.isfile(filename):
        os.remove(filename)
    print ("Report: ", filename)

    # Wait for Excel to launch
    print ("Sleeping while Excel loads...")
    time.sleep (90)

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
print ("Done!")
