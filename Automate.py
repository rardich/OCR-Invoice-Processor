import pyautogui, csv

# Failsafe by moving mouse to upper left hand corner of display
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.1

width, height = pyautogui.size()

print('Press Ctrl + C to quit program')


# Need to implement on actual computer that will be used to interface with Epicor
try:
    while True:
        pyautogui.moveTo(100, 100, duration = 0.25)
        pyautogui.moveTo(100, 200, duration = 0.25)
        pyautogui.moveTo(200, 200, duration = 0.25)
        pyautogui.moveTo(200, 100, duration = 0.25)
except KeyboardInterrupt:
    print('Exited Sucessfully')