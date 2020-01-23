"""
This script will open.par files and combine them into a single, multisheet excel file"""

import pandas as pd
import glob
import os
import openpyxl

#This retrieves the path of the script and searches for all .par files

abspath = os.path.abspath("XPSParSummary 1.2.py")
path = os.path.dirname(abspath)

filenames1 = glob.glob(path + "/*.par")
filenames2 = [os.path.split(y)[1] for y in filenames1]
filenames = [os.path.splitext(os.path.basename(z))[0] for z in filenames2]

#Finds and retrieves .xls file to use as FILENAME - if no .xls will simply name the files 'sample'
core_file_extpath = glob.glob(path + "/*.xls")
core_fileext = [os.path.split(h)[1] for h in core_file_extpath]
core_file = [os.path.splitext(os.path.basename(j))[0] for j in core_fileext]
if len(core_file) != 0:
    core_file_s = core_file[0]
else:
    core_file_s = input("Excel file not found - please manually type sample name:")

#Output excel file is created

writer = pd.ExcelWriter(core_file_s+" summary.xlsx")

for file in filenames:

    title = str(file)+" region in "+str(core_file_s)
    sn = pd.read_fwf(file+".par", delim_whitespace=True)

    #Confirms file read
    print("Collected "+file)

    #calculates percentage integration and confirms totals add up to 100%
    columns = (list(sn.columns))
    ints = list(sn[columns[3]])
    int_percents = []
    total_int = sum(ints)
    for integration in ints:
        int_percents.append(round(((integration/total_int)*100), 4))
    sn["Percent Integration"] = int_percents
    
    #Total Integration Percentage Sum should be 100 as a sanity check
    sn["Total Integration Percentage Sum"] = sum(int_percents)

    #writes file to excel sheet - separate sheet for each .par file
    sn.to_excel(writer, sheet_name=file)

#finally saves file
writer.save()
print(core_file_s+" summary.xlsx has been created successfully.")
