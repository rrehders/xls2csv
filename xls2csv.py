#! /usr/bin/env python3

import sys
import os
import openpyxl
import csv

__author__ = 'rrehders'

# Validate the correct number of commanf line args
if not len(sys.argv) == 2:
    print('USAGE: XLS2CSV [file1]')
    sys.exit()

# Validate command line arg is a file
fname = sys.argv[1]
if not os.path.isfile(fname):
    print('ERR: '+fname+' is not a file')
    sys.exit()

# load the target workbook
try:
    wb = openpyxl.load_workbook(fname, data_only=True)
except Exception as err:
    print('ERR: '+fname+' '+str(err))

# Display Sheets in the workbook and ask which sheet to convert
# A CSV can only contain a single sheet
print('XLS2CSV: Convert an Excel worksheet to CSV')
print('Sheets in '+fname)
sheetnms = wb.get_sheet_names()
for i in range(len(sheetnms)):
    print(' | '+str(i)+' - '+sheetnms[i], end='')
print(' |')

# Get sheet selection
sheetnum = -1
while sheetnum not in range(len(sheetnms)):
    sheetnum = int(input('Convert which sheet ? '))
print('')

# Set the active sheet to the selection
xlsheet = wb.get_sheet_by_name(sheetnms[sheetnum])

# Build lists of values for each row
print('Extracting values')
table = []
for rowOfCellObjs in xlsheet:
    row = []
    for cellObj in rowOfCellObjs:
        row += [cellObj.value]
        print('.', end='')
    table += [row]
    print('')

# Set the output filename based on sheet name
ofname = sheetnms[sheetnum]+'.csv'

# Open output file
ofile = open(ofname, mode='w', newline='')

# attach csv_writer to the output file
oWriter = csv.writer(ofile)

# Write out each row of the csv file
print('Writing file '+ofname+'.')
for i in range(len(table)):
    print('.', end='')
    oWriter.writerow(table[i])
print('')

# Close output file
ofile.close()
