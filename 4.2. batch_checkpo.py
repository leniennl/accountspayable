import pyautogui
import ctypes
import os
from pynput import keyboard
import pandas as pd
import shutil

source_folder = r"C:\Users\matthew.lee\Invoice\\"
destination_folder = r"C:\Users\matthew.lee\To Send to Branches\\"

df = pd.read_csv("invoicedata.csv")

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 1


def on_press(key):
    global status_invoice
    global filename
    status_invoice = ""
    if key == keyboard.Key.esc:
        df.to_csv("invoicedata.csv", index=False)  # save invoice status before escaping
        print("quit script " + k)
        quit()
        return False  # stop listener
    try:
        k = key.char  # single-char keys
    except:
        k = key.name  # other keys
    if k in ["left", "right", "space"]:  # keys of interest
        # self.keys.append(k)  # store it in global-like variable
        df.to_csv("invoicedata.csv", index=False)
        print("quit script " + k)
        quit()
        return False  # stop listener; remove this if want more keys
    if k in ["m", "M"]:
        print("to match")
        status_invoice = "to match"
        return False
    if k in ["r", "R"]:
        print("to receipt")
        status_invoice = "to receipt"
        source = source_folder + filename
        destination = destination_folder + filename
        shutil.move(source, destination)
        print("Moved:", filename)
        return False
    if k in ["a", "A"]:
        print("to amend")
        status_invoice = "to amend"
        source = source_folder + filename
        destination = destination_folder + filename
        shutil.move(source, destination)
        print("Moved:", filename)
        return False


# close invoicedata.csv if it's already opened
while True:  # repeat until the try statement succeeds   # or "a+", whatever you need
    try:
        os.rename("invoicedata.csv", "invoicedata.csv")
        break  # exit the loop
    except IOError:
        input("Please Close Invoicedata.csv")
        quit()


for subdir, dirs, files in os.walk(r"C:\Users\matthew.lee\Invoice"):
    for filename in files:
        if filename.endswith(".pdf"):
            po = filename.split(" ")[1]
            invoice_number = filename.split(" ")[3]
            supplier = filename.split(" ")[5].split(".")[0]
            ctypes.windll.user32.SetCursorPos(650, 212)  # cancel previous po entry
            pyautogui.click()
            ctypes.windll.user32.SetCursorPos(710, 212)  # cancel previous po entry
            pyautogui.click()
            print(
                "Now checking "
                + str(supplier)
                + "      PO: "
                + str(po)
                + "       INV: "
                + str(invoice_number)
            )
            ctypes.windll.user32.SetCursorPos(620, 220)  # enter PO
            pyautogui.click()
            pyautogui.write(po, interval=0.3)
            ctypes.windll.user32.SetCursorPos(350, 170)  # find
            pyautogui.click()
            ctypes.windll.user32.SetCursorPos(300, 170)  # select
            pyautogui.click()
            ctypes.windll.user32.SetCursorPos(455, 575)  # go to end
            pyautogui.click()
            ctypes.windll.user32.SetCursorPos(1860, 570)  # move to max
            pyautogui.click()
            listener = keyboard.Listener(on_press=on_press)
            listener.start()  # start to listen on a separate thread
            listener.join()  # remove if main thread is polling self.keys
            df.loc[df.invoice_number == invoice_number, "inv_status"] = status_invoice
            ctypes.windll.user32.SetCursorPos(360, 175)  # cancel po enquiry
            pyautogui.click()

df.to_csv("invoicedata.csv", index=False)
print("All PO checked in the Invoice folder.")

# move files from Invoice to To Send to Branches


# # fetch all files
# for file_name in os.listdir(source_folder):
#     # construct full file path
#     source = source_folder + file_name
#     destination = destination_folder + file_name
#     # move only files
#     if os.path.isfile(source):
#         shutil.move(source, destination)
#         print('Moved:', file_name)
