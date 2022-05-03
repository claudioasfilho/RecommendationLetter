
#import ListofThings
import xlrd
from xlrd import open_workbook
import sys, getopt
import os

#it reads the input file
StudentProfile_xls = "StudentProfile.xls"

if os.path.isfile(StudentProfile_xls):
    #it Imports the Excel File
    wb = open_workbook(StudentProfile_xls)
    ws= wb.sheet_by_index(0)
    number_of_rows = ws.nrows
    number_of_columns = ws.ncols
    firstName = (ws.cell(3,1).value)
    lastName = (ws.cell(4,1).value)

else:
    print("{0} does not appear to be a valid file, choose the right file and retry".format(inputfile))
    sys.exit(2)

print("first Name is " + firstName + lastName)
print("Last Name is " + lastName)
