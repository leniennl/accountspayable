import pyautogui
import ctypes
import pandas as pd

pyautogui.FAILSAFE = True

# prompt for invoice number
invoicenumber = input("Please enter invoice number:\n")
invoicenumber = invoicenumber.capitalize().strip()


# use invoice number to fetch PO,gross,invoice_date
df = pd.read_csv("invoicedata.csv")
if invoicenumber not in df.invoice_number.values:
    print("Invoice Number Not Registered.\n")
    exit()
vendor = df.loc[df.invoice_number == invoicenumber, "company"].to_string(index=False)
po = df.loc[df.invoice_number == invoicenumber, "PO"].to_string(index=False)
gross = df.loc[df.invoice_number == invoicenumber, "invoice_total"].to_string(
    index=False
)
invoicedate = df.loc[df.invoice_number == invoicenumber, "invoice_date"].to_string(
    index=False
)
if "/" in invoicedate:
    invoicedate = (
        invoicedate.split("/")[1].rjust(2, "0")
        + invoicedate.split("/")[0].rjust(2, "0")
        + invoicedate.split("/")[2][2:4]
    )
elif "-" in invoicedate:
    invoicedate = (
        invoicedate.split("-")[2][0:2]
        + invoicedate.split("-")[1]
        + invoicedate.split("-")[0][2:4]
    )


# auto fill information in Voucher Match screen
"""po"""
ctypes.windll.user32.SetCursorPos(600, 210)
pyautogui.click()
pyautogui.write(" " + po, interval=0.2)
ctypes.windll.user32.SetCursorPos(920, 300)
pyautogui.click()
"""invoice number"""
ctypes.windll.user32.SetCursorPos(600, 300)
pyautogui.click()
pyautogui.write(invoicenumber, interval=0.2)
ctypes.windll.user32.SetCursorPos(920, 300)
pyautogui.click()
"""gross"""
ctypes.windll.user32.SetCursorPos(660, 330)
pyautogui.click()
pyautogui.write(gross, interval=0.2)
ctypes.windll.user32.SetCursorPos(920, 300)
pyautogui.click()
"""invoicedate"""
ctypes.windll.user32.SetCursorPos(660, 366)
pyautogui.click()
pyautogui.write(invoicedate, interval=0.2)
pyautogui.hotkey("tab")
"""Co."""
ctypes.windll.user32.SetCursorPos(1000, 300)
pyautogui.click()
pyautogui.write("112", interval=0.2)


# form open receipt
ctypes.windll.user32.SetCursorPos(455, 175)
pyautogui.click()
pyautogui.PAUSE = 1.5
ctypes.windll.user32.SetCursorPos(460, 255)
pyautogui.click()
ctypes.windll.user32.SetCursorPos(450, 385)
pyautogui.click()  # goto end
if vendor != "ThermoFisher":
    ctypes.windll.user32.SetCursorPos(315, 455)
    pyautogui.click()  # select all
    ctypes.windll.user32.SetCursorPos(300, 170)
    pyautogui.click()  # ok


# mark excel records as completed
