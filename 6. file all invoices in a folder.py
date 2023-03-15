import glob
import os
import shutil

src_folder = r"C:\Users\matthew.lee\To Send to Branches"

thermofisher_folder = r"W:\Accounts Payable\Thermofisher\\"
symbio_folder = r"W:\Accounts Payable\Symbio Laboratories\\"
als_folder = r"W:\Accounts Payable\ALS Environmental\\"
nmi_folder = r"W:\Accounts Payable\NMI\\"

# move file whose name with string
pattern = src_folder + "\*ThermoFisher*"
for file in glob.iglob(pattern, recursive=True):
    # extract file name form file path
    file_name = os.path.basename(file)
    shutil.move(file, thermofisher_folder + file_name)
    print("Filed:", file)

# move file whose name with string
pattern = src_folder + "\*ALS*"
for file in glob.iglob(pattern, recursive=True):
    # extract file name form file path
    file_name = os.path.basename(file)
    shutil.move(file, als_folder + file_name)
    print("Filed:", file)

# move file whose name with string
pattern = src_folder + "\*nmi*"
for file in glob.iglob(pattern, recursive=True):
    # extract file name form file path
    file_name = os.path.basename(file)
    shutil.move(file, nmi_folder + file_name)
    print("Filed:", file)

# move file whose name with string
pattern = src_folder + "\*symbio*"
for file in glob.iglob(pattern, recursive=True):
    # extract file name form file path
    file_name = os.path.basename(file)
    shutil.move(file, symbio_folder + file_name)
    print("Filed:", file)
