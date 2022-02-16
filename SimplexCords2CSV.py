import keyboard  # using module keyboard
import time
import pyperclip
import xlsxwriter
import pyautogui
import time
from win32gui import GetWindowText, GetForegroundWindow
lastdata = 0
csvlist = []
pyperclip.copy('A')
try:
    while True:  # making a loop
        if('Simplex Mapping' in GetWindowText(GetForegroundWindow()) or 'BuildTheEarth' in GetWindowText(GetForegroundWindow())):
            data = pyperclip.paste()
            if(data != lastdata):
                lastdata = data
                #34.7740972, 32.0983703, 2.66
                print(data.count(','))
                if(data.find(',') != -1 and data.count(',') == 2 and '.' in data and data.count('.') == 3):
                    print("Copied coordinates from Simplex: " + data)
                    text = data.split(',')
                    if(len(text) > 2):
                        command = '/tpll ' + text[1].lstrip(' ') + " " + text[0] + " " + str(round(float(text[-1])))
                        print("Clipboard updated with: " + command)
                        print(GetWindowText(GetForegroundWindow()))
                        csvlist.append(command)
        if(keyboard.is_pressed('ctrl + shift + a')):
            break
    for z in range(len(csvlist)):
        print(csvlist[z])
    workbook = xlsxwriter.Workbook('Coordinates.xlsx')
    worksheet = workbook.add_worksheet()
    for val in range(len(csvlist)):
        worksheet.write_column('A' + str(val), csvlist[val])
    workbook.close()    
    print(GetWindowText(GetForegroundWindow()))
    print(data)
    selection = input("What type of selection? (Convex,Poly)")
    input("Press Enter to continue...")
    numoffile = 0
    while(GetWindowText(GetForegroundWindow()) != 'Minecraft 1.12.2 - BuildTheEarth'):
        time.sleep(1)
    print("Staring in 1")
    time.sleep(1)
    pyautogui.press('t')
    pyautogui.write('//sel ' + selection)
    pyautogui.press('enter')
    pyautogui.press('t')
    pyautogui.write('//sel')
    pyautogui.press('enter')
    time.sleep(.5)
    for x in csvlist:
        if(GetWindowText(GetForegroundWindow()) != 'Minecraft 1.12.2 - BuildTheEarth'):
            quit()
            exit()
        else:
            numoffile += 1
            print(str(numoffile) + '/' + str(len(csvlist)))
            pyautogui.press('t')
            pyautogui.write(str(x))
            pyautogui.press('enter')
            time.sleep(.5)
            pyautogui.press('t')
            if(x == csvlist[0]):
                pyautogui.write("//pos1")
            else:
                pyautogui.write("//pos2")
            pyautogui.press('enter')
except KeyboardInterrupt:
    print("Ended")

