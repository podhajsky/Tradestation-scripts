'''
(c) 2016, Petr Podhajsky, www.financnik.cz.
Script for copying P/L History from TradeStation Walk-Forward Optimizer into text files

!!! Script has to be run under Admin rights
!!! After start click focus on Tradestation Walk-Forward Optimizer window
'''

import pyautogui
import time
import pandas as pd
import win32clipboard
from io import StringIO

excel_filename = "pl_history.xls"

# wait 5 sec so we have time to manually switch focus to TS WFO window
print ('Click on TradeStation Walk-Forward Optimizer windows. Waiting 5 sec.')
time.sleep(5)

# first find position of open WFO button, to get some coordinates.
# it is possible you have to do your own screenshot
button_loc = pyautogui.locateOnScreen('images/open-WFO-button.png')

#definition of OOS/Run matrix. Runs need to be string sorted alphabetically to match TradeStation sorting.
runs = [str(x) for x in range(5, 35, 5)]
runs.sort()
ooss = [str(x) for x in range(10, 35, 5)]

# we can continue  only if we have coordinate:
if button_loc is not None:

    print("Recevied location of TS WFO window. Start processing....")

    #prepare excel file so we can store sheets
    writer = pd.ExcelWriter(excel_filename)

    #main loop - iterate over OOS and RUNS:
    for oos in ooss:
        for run in runs:
            # click to bin selections
            pyautogui.click(button_loc[0] + 107, button_loc[1] + 45)

            # in the first run go up in the list
            if (oos == '10') and (run == '10'):
                # press up several times to make sure we get on first position
                for x in range(0, 31):
                    pyautogui.press('up')
                # then press enter to select bin
                pyautogui.press('enter')
            else:
                pyautogui.press('down')
                pyautogui.press('enter')

            # create bin name so we can identify data
            bin_name = "OOS" + oos + "% " + "WFRuns=" + run

            print ('Collecting data for: ', bin_name)

            # click on "P/L history"
            pyautogui.click(button_loc[0] + 608, button_loc[1] + 75)

            # click inside trade list
            pyautogui.click(button_loc[0] + 264, button_loc[1] + 386)

            # press CTRL + A to select all
            pyautogui.hotkey('ctrl', 'a')
            # copy to clipboard
            pyautogui.hotkey('ctrl', 'c')

            # get data from clipboard into variable:
            win32clipboard.OpenClipboard()
            clipboard_data = win32clipboard.GetClipboardData()
            win32clipboard.CloseClipboard()

            # parse data from clipboard as csv file. Use Tab as separator, rename columns
            df = pd.read_csv(StringIO(clipboard_data), sep="\t", index_col=0, header=None,
                             names=['Exit Date', 'Exit Time', 'Position', 'Shares/Ctrts',
                                    'Net Profit', 'Cum Net Prft', 'Drawdown', 'Bars'])

            # TS print data as string. Convert string values into floats:
            df['Net Profit'] = df['Net Profit'].map(lambda x: float(str(x).replace(',', '')))
            df['Cum Net Prft'] = df['Cum Net Prft'].map(lambda x: float(str(x).replace(',', '')))
            df['Drawdown'] = df['Drawdown'].map(lambda x: float(str(x).replace(',', '')))

            # store the run into excel sheet
            df.to_excel(writer, sheet_name=bin_name)

    #save excel file
    print ('Saving excel file as: ', excel_filename)
    writer.save()


else:
    print ("Could not locate the button to get coordinates. End.")