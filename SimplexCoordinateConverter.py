import keyboard  # using module keyboard
import time
import pyperclip
from win32gui import GetWindowText, GetForegroundWindow
import pyautogui
lastdata = 0
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
                        command = '/tpll ' + text[1].lstrip(' ') + " " + text[0] + " " + str(round(float(text[2])))
                        print("Clipboard updated with: " + command)
                        print(GetWindowText(GetForegroundWindow()))
                        while(pyperclip.paste() != command):
                            pyperclip.copy(command)
                            print("Command not copied.")
                        print("Copied: " + command)   
        if(keyboard.is_pressed('ctrl + shift + a')):
            print(command)
            pyautogui.write(command)
except KeyboardInterrupt:
    print("Ended")

