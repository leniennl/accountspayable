import os
import invoice2data
from invoice2data import extract_data
from invoice2data.extract.loader import read_templates
import csv
import datetime
import xlrd
import openpyxl
import docx
import re

os.system("cls")

PASSWORD = "Silliker7"


def getText(filename):
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    return "\n".join(fullText)


def turnDTintoSTR(somedatetimeobject):
    if isinstance(somedatetimeobject, list):
        newlistwithoutnone = [i for i in somedatetimeobject if i is not None]
        if len(newlistwithoutnone) >= 1:
            somedatetimeobject = min(newlistwithoutnone)
            somedatetimeobject = str(
                str(somedatetimeobject.day).rjust(2, "0")
                + str(somedatetimeobject.month).rjust(2, "0")
                + str(somedatetimeobject.year)[2:4]
            )
        else:
            somedatetimeobject = ""
    elif isinstance(somedatetimeobject, datetime.date):
        somedatetimeobject = str(
            str(somedatetimeobject.day).rjust(2, "0")
            + str(somedatetimeobject.month).rjust(2, "0")
            + str(somedatetimeobject.year)[2:4]
        )
    return somedatetimeobject


def scrapePDFAuditorInvoice(filepath):
    # templates = read_templates('Invoice2data/Template/')
    templates = read_templates(
        r"C:\Users\matthew.lee\Dropbox\Side Hussle\Python\Finished Projects\Invoice2data\Template"
    )
    today = datetime.datetime.now()
    today = str(today.day).zfill(2) + str(today.month).zfill(2) + str(today.year)
    result = extract_data(filepath, templates=templates)
    result["register_date"] = today  # add time stamp
    result["date"] = str(turnDTintoSTR(result["date"]))

    with open("auditorinvoicedata.csv", "a", newline="") as f:
        w = csv.writer(f)
        w.writerow(result.values())

    os.rename(
        filepath,
        "Invoice/" + str(result["issuer"] + " " + result["invoice_number"] + ".pdf"),
    )


def scrapeDOCXAuditorInvoice(docxfilepath):
    whole_docx_text = getText(docxfilepath)

    # build re object
    vendor_name_re = re.compile(r"(KAR FOOD SAFETY AUDITING PTY LTD)")
    invoice_number_re = re.compile(r"Invoice No:\s+(\d+)")
    invoice_total_re = re.compile(r"\nTotal\s+\$\s+(\d+\.\d{2})")
    invoice_date_re = re.compile(r"Date:\s+(\d{1,2}/\d{2}/\d{2,4})")

    # search re object
    if vendor_name_re.search(whole_docx_text) is not None:
        vendorname = vendor_name_re.search(whole_docx_text).group(1)
        vendornumber = 504900
    else:
        vendorname = ""
        vendornumber = ""

    if invoice_number_re.search(whole_docx_text) is not None:
        invoicenumber = invoice_number_re.search(whole_docx_text).group(1)
    else:
        invoicenumber = ""

    if invoice_total_re.search(whole_docx_text) is not None:
        invoicetotal = invoice_total_re.search(whole_docx_text).group(1)
    else:
        invoicetotal = ""

    if invoice_date_re.search(whole_docx_text) is not None:
        invoicedate = invoice_date_re.search(whole_docx_text).group(1)
    else:
        invoicedate = ""

    today = datetime.datetime.now()
    today = str(today.day).zfill(2) + str(today.month).zfill(2) + str(today.year)
    outputdict = {
        "Vendor Name": vendorname,
        "Vendor Number": vendornumber,
        "Inv Number": invoicenumber,
        "Inv Total": invoicetotal,
        "Inv Date": invoicedate,
        "Currency": "AUD",
        "Desc": "Docx",
        "Process Date": today,
    }

    with open(r"C:\Users\matthew.lee\auditorinvoicedata.csv", "a", newline="") as f:
        w = csv.writer(f)
        w.writerow(outputdict.values())

    os.rename(
        docxfilepath , "Invoice/" + str(vendorname) + " " + str(invoicenumber) + ".docx"
    )


def scrapeXLSAuditorInvoice(filename):
    wb = xlrd.open_workbook("C:/Users/matthew.lee/invoice/" + filename)
    for sh in wb.sheets():
        if sh.name == "Auditor Invoice":
            for row in range(sh.nrows):
                for col in range(sh.ncols):
                    myCell = sh.cell(row, col)
                    if myCell.value == "Invoice Total":
                        invoicetotal = sh.cell(row, col + 3).value
                        invoicetotal = round(invoicetotal, 2)
                    elif myCell.value == "Inv No":
                        invoicenumber = sh.cell(row, col + 1).value
                    elif myCell.value == "Inv Date":
                        invoicedate = sh.cell(row, col + 1).value
                        if type(invoicedate) is float:
                            invoicedate = xlrd.xldate_as_datetime(invoicedate, 0)
                            invoicedate = invoicedate.date()
                    elif myCell.value == "Vendor Number":
                        vendornumber = sh.cell(row, col + 1).value
                    elif myCell.value == "Auditor Company Name":
                        vendorname = sh.cell(row, col + 1).value

            today = datetime.datetime.now()
            today = (
                str(today.day).zfill(2) + str(today.month).zfill(2) + str(today.year)
            )
            outputdict = {
                "Vendor Name": vendorname,
                "Vendor Number": vendornumber,
                "Inv Number": invoicenumber,
                "Inv Total": invoicetotal,
                "Inv Date": invoicedate,
                "Currency": "AUD",
                "Desc": "xls",
                "Process Date": today,
            }

            with open(
                r"C:\Users\matthew.lee\auditorinvoicedata.csv", "a", newline=""
            ) as f:
                w = csv.writer(f)
                w.writerow(outputdict.values())

            os.rename(
                "Invoice/"+filename, "Invoice/" + str(vendorname).replace("/","") + " " + str(invoicenumber) + ".xls"
            )


def scrapeXLSXAuditorInvoice(wb):
    wrb = openpyxl.open(wb, read_only=True, data_only=True)
    sh = wrb["Auditor invoice"]
    for row in range(1, sh.max_row + 1):
        for col in range(1, sh.max_column + 1):
            myCell = sh.cell(row, col)
            if myCell.value == "Invoice Total":
                invoicetotal = sh.cell(row, col + 2).value
            elif myCell.value == "Inv No":
                invoicenumber = sh.cell(row, col + 1).value
            elif myCell.value == "Inv Date":
                invoicedate = sh.cell(row, col + 1).value
                invoicedate = invoicedate.strftime("%d%m%y")
            elif myCell.value == "3532-13706":
                vendornumber = 506930
            elif myCell.value == "Auditor Company Name":
                vendorname = sh.cell(row, col + 1).value

    today = datetime.datetime.now()
    today = str(today.day).zfill(2) + str(today.month).zfill(2) + str(today.year)
    outputdict = {
        "Vendor Name": vendorname,
        "Vendor Number": vendornumber,
        "Inv Number": invoicenumber,
        "Inv Total": invoicetotal,
        "Inv Date": invoicedate,
        "Currency": "AUD",
        "Desc": "xlsx",
        "Process Date": today,
    }

    with open(r"C:\Users\matthew.lee\auditorinvoicedata.csv", "a", newline="") as f:
        w = csv.writer(f)
        w.writerow(outputdict.values())

    os.rename(wb, "Invoice/" + str(vendorname).replace("/","") + " " + str(invoicenumber) + ".xlsx")


def walkfiles():
    unprocessed_files = []
    for subdir, dirs, files in os.walk(r"C:\Users\matthew.lee\invoice"):
        for filename in files:
            if filename.endswith(".xls"):             
                try:
                    scrapeXLSAuditorInvoice(filename)
                except:
                    unprocessed_files.append(filename)
                    continue
            elif filename.endswith(".xlsx"):
                try:
                    xlsxwb = "Invoice/" + filename
                    scrapeXLSXAuditorInvoice(xlsxwb)
                except:
                    unprocessed_files.append(filename)
                    continue
            elif filename.endswith(".pdf"):
                try:
                    path = "Invoice/" + filename
                    scrapePDFAuditorInvoice(path)
                except:
                    unprocessed_files.append(filename)
                    continue
            elif filename.endswith(".docx"):
                try:
                    docxpath = "Invoice/" + filename
                    scrapeDOCXAuditorInvoice(docxpath)
                except:
                    unprocessed_files.append(filename)
                    continue
            else:
                unprocessed_files.append(filename)
                continue
    if len(unprocessed_files) > 0:
        print("\n\nThe Following Invoices are Not Processed:\n")
        print("\n".join(map(str, unprocessed_files)))
    else:
        print("\n\nAll Invoices Info Extracted!\n")


# close invoicedata.csv if it's already opened
try:
    os.rename("auditorinvoicedata.csv", "auditorinvoicedata.csv")
except IOError:
    input("Please Close auditornvoicedata.csv")
    quit()
else:
    walkfiles()
