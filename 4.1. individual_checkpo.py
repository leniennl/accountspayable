import pyautogui
import ctypes

# cursor = ctypes.wintypes.POINT()
# print(ctypes.windll.user32.GetCursorPos(ctypes.byref(cursor)))

# print(ctypes.windll.user32.GetCursorPos(ctypes.byref(cursor)))
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 1

ctypes.windll.user32.SetCursorPos(350, 170)  # find
pyautogui.click()
ctypes.windll.user32.SetCursorPos(300, 170)  # select
pyautogui.click()
ctypes.windll.user32.SetCursorPos(455, 575)  # go to end
pyautogui.click()
ctypes.windll.user32.SetCursorPos(1860, 570)  # move to max
pyautogui.click()
