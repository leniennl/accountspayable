import os
import invoice2data
from invoice2data import extract_data
from invoice2data.extract.loader import read_templates
import csv
import datetime as dt

# close invoicedata.csv if it's already opened
while True:  # repeat until the try statement succeeds   # or "a+", whatever you need
    try:
        os.rename("invoicedata.csv", "invoicedata.csv")
        break  # exit the loop
    except IOError:
        input("Please Close Invoicedata.csv")
        quit()
        # os.system('taskkill /F /IM excel.exe')
        # time.sleep(2)
        # restart the loop


templates = read_templates("Invoice2data/Template/")
today = dt.datetime.now()
today = str(today.day).zfill(2) + str(today.month).zfill(2) + str(today.year)

for subdir, dirs, files in os.walk(r"C:\Users\matthew.lee\invoice"):
    for filename in files:
        if filename.endswith(".pdf"):

            # read in template from invoice2data and extract invoice information
            result = extract_data("Invoice/" + filename, templates=templates)

            # standardize state section:
            if result.get("account_number") != "":
                state = str(result.get("account_number"))
                state = state.upper()
                if state == "228980" or state == "SILL02" or state[0] == "2":
                    state = "NSW"
                elif (
                    state == "305570"
                    or state == "SILL03"
                    or state[0] == "3"
                    or state == "HEA"
                ):
                    state = "VIC"
                elif state == "401562" or state == "SILL05" or state[0] == "4":
                    state = "QLD"
                elif state == "623296" or state == "SILL04" or state[0] == "6":
                    state = "WA"
                result["account_number"] = state

            # add time stamp
            result["register_date"] = today

            # rename invoice as per state_supplier_PO_invoice number
            newfname = (
                "PO "
                + str(result.get("purchase_order"))
                + " INV "
                + str(result.get("invoice_number"))
                + " "
                + str(state)
                + " "
                + str(result.get("company"))
                + ".pdf"
            )
            oldfpath = subdir + os.sep + filename
            newfpath = subdir + os.sep + newfname
            os.rename(oldfpath, newfpath)

            # acknowledgement
            print(f"Changed file name from {filename} to {newfname}. ")

            # # write dictionary to Excel for future processing
            # wb=xw.Book('invoicedata.xlsm')
            # finalrow=wb.sheets[0].range('A' + str(wb.sheets[0].cells.last_cell.row)).end('up').row
            # dict=pd.DataFrame(data=result,index=[0])
            # wb.sheets[0].range('A' + str(finalrow+1)).options(index=False,header=False).value=dict

            # try CSV approach of saving invoice info data
            with open("invoicedata.csv", "a", newline="") as f:
                w = csv.writer(f)
                w.writerow(result.values())

            # try pandas to_excel approach
            # pd.DataFrame(data=result,index=[0]).to_excel("invoicedata.xlsx")

            print("Invoice info saved.")
