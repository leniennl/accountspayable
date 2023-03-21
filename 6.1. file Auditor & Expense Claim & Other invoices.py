import glob
import os
import shutil

src_folder = r"C:\Users\matthew.lee\Invoice\\"

auditor_invoice_folder = r"W:\Accounts Payable\Auditors Invoices\\"

employee_expense_folder = r"W:\Accounts Payable\Employee Expense Claims\\"


# move file whose name with string
for root,dirs,files in os.walk(src_folder):
    for name in files:
        if "Claim Form" not in name:
            shutil.move(src_folder + name, auditor_invoice_folder + name)
            print("Filed:", name)
        else:
            shutil.move(src_folder + name, employee_expense_folder + name)
            print("Filed:", name)

