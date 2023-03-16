import pyautogui
import ctypes
import pandas as pd
import datetime as dt
import time
from pynput import keyboard, mouse
import sys

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.1

df = pd.read_csv(r"C:\Users\matthew.lee\auditorinvoicedata.csv")
df = df.drop_duplicates()

EXPECTED_COLOR = (58, 58, 58)


def on_press(key):
    try:
        k = key.char  # single-char keys
    except:
        k = key.name  # other keys
    if k in ["g", "G"]:  # Proceed to Save
        return False  # stop listener
    if k in ["right", "left", "up", "down", "enter"]:
        return False
    if k in ["esc"]:
        print("quit script ")
        quit()
        return False  # stop listener


def on_click(x, y, button, pressed):
    if pressed and pyautogui.pixelMatchesColor(x, y, EXPECTED_COLOR):
        print(f"Mouse clicked at position: ({x}, {y})")
        return False



"""
# Start listening for mouse click events
print("Click top edge of terminal window.")
with mouse.Listener(on_click=on_click) as listener:
    listener.join()
# drag terminal
pyautogui.click()
pyautogui.dragTo(2318, 131, 3, button="left")
"""


def processDate(somedate):
    somedate = str(somedate)
    if "/" in somedate:
        if len(somedate.split("/")[2]) == 4:
            somedate = (
                somedate.split("/")[0].rjust(2, "0")
                + somedate.split("/")[1].rjust(2, "0")
                + somedate.split("/")[2][2:4]
            )
        elif len(somedate.split("/")[2]) == 2:
            somedate = (
                somedate.split("/")[0].rjust(2, "0")
                + somedate.split("/")[1].rjust(2, "0")
                + somedate.split("/")[2]
            )
    elif "-" in somedate:
        somedate = (
            somedate.split("-")[2][0:2]
            + somedate.split("-")[1]
            + somedate.split("-")[0][2:4]
        )
    elif "." in somedate:
        somedate1 = somedate.split(".")[0]
        somedate2 = somedate.split(".")[1]
        somedate3 = somedate.split(".")[2]
        if len(somedate1) == 1:
            somedate1 = somedate1.rjust(2, "0")
        if len(somedate2) == 1:
            somedate2 = somedate2.rjust(2, "0")
        if len(somedate3) == 4:
            somedate3 = somedate3[2:4]
        somedate = somedate1 + somedate2 + somedate3
    elif " " in somedate:
        somedate = somedate.strip()
    if len(somedate) < 6:
        somedate = "0" + somedate
    return somedate


def main():
    # reposition terminal
    fw = pyautogui.getWindowsWithTitle("cmd.exe")
    if fw[0].topleft != (2318, 281):
        fw[0].topleft = (2318, 281)
    fw[0].activate()
    # prompt for invoice number
    invoicenumber = input("Please enter invoice number:\n")
    invoicenumber = invoicenumber.strip()
    while invoicenumber not in df.invoice_number.values:
        if invoicenumber.lower() == "bye":
            sys.exit()
        print("Invoice Number Not Registered.\n")
        invoicenumber = input("Please enter invoice number:\n")
        invoicenumber = invoicenumber.strip()

    # use invoice number to fetch vendor_number,gross,invoice_date,GL date
    vendor = df.loc[df.invoice_number == invoicenumber, "company"].to_string(
        index=False
    )
    vendornumber = df.loc[
        df.invoice_number == invoicenumber, "vendor_number"
    ].to_string(index=False)
    gross = df.loc[df.invoice_number == invoicenumber, "invoice_total"].to_string(
        index=False
    )
    invoicedate = df.loc[df.invoice_number == invoicenumber, "invoice_date"].to_string(
        index=False
    )

    invoicedate = processDate(invoicedate)
    today = dt.datetime.now()
    gldate = str(today.day).zfill(2) + str(today.month).zfill(2) + str(today.year)[2:]
    gldate = processDate(gldate)
    remark = " AUDITORS"

    # auto fill information in Voucher Match screen
    """bring JDE to focus"""
    ctypes.windll.user32.SetCursorPos(610, 60)
    pyautogui.click()
    """company"""
    ctypes.windll.user32.SetCursorPos(610, 260)
    pyautogui.click()
    pyautogui.write(" 112", interval=0.5)
    ctypes.windll.user32.SetCursorPos(1000, 300)
    pyautogui.click()

    """vendor number"""
    ctypes.windll.user32.SetCursorPos(800, 300)
    pyautogui.click()
    pyautogui.write(" " + vendornumber, interval=0.5)
    ctypes.windll.user32.SetCursorPos(1000, 300)
    pyautogui.click()

    """invoice number"""
    ctypes.windll.user32.SetCursorPos(610, 340)
    pyautogui.click()
    pyautogui.write(invoicenumber, interval=0.5)
    ctypes.windll.user32.SetCursorPos(1000, 300)
    pyautogui.click()

    """invoice date"""
    ctypes.windll.user32.SetCursorPos(610, 370)
    pyautogui.click()
    pyautogui.write(invoicedate, interval=0.5)
    pyautogui.press("tab")

    """GL date"""
    ctypes.windll.user32.SetCursorPos(1050, 370)
    pyautogui.click()
    pyautogui.write(gldate, interval=0.5)
    pyautogui.press("tab")

    """gross"""
    ctypes.windll.user32.SetCursorPos(510, 530)
    pyautogui.click()
    pyautogui.write(" " + gross, interval=0.5)

    """remark"""
    ctypes.windll.user32.SetCursorPos(850, 530)
    pyautogui.click()
    pyautogui.write(" " + remark, interval=0.5)

    listener = keyboard.Listener(on_press=on_press)
    listener.start()  # start to listen on a separate thread
    listener.join()  # check invoice before saving
    """save invoice entry"""
    ctypes.windll.user32.SetCursorPos(310, 175)
    pyautogui.click()

    time.sleep(3)
    """save invoice entry AGAIN"""
    ctypes.windll.user32.SetCursorPos(310, 175)
    pyautogui.click()
    time.sleep(1)

    # reactivate terminal window for next input
    fw[0].activate()
    """
    ctypes.windll.user32.SetCursorPos(2318, 131)
    pyautogui.click()
    """


while True:  # repeat until the try statement succeeds   # or "a+", whatever you need
    try:
        main()
    except IOError:
        quit()
