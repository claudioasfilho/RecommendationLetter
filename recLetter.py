
import Listofthings
import openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook
import sys, getopt
import os

#it reads the input file
StudentProfile_xls = "StudentProfile.xlsx"


if os.path.isfile(StudentProfile_xls):
    #it Imports the Excel File
    wb = load_workbook(StudentProfile_xls)
    sheets = wb.sheetnames
    ws= wb[sheets[0]]

    firstName = ws['B3'].value
    lastName = ws['B4'].value
    pronoum = ws['B5'].value
    classAttended = ws['B6'].value
    schoolYearAttended = ws['B7'].value
    highSchoolYearAttended = ws['B8'].value

    accomplishments = 0
    #Column A
    colA = ws['A']
    #Column B
    colB = ws['B']


    ############################
    #Collecting purposeOfTheLetter
    ############################
    dataAvailable = 0

    #Initial Row where data will be examined, The offset needs to be set at -1
    rowOffSet = 12 - 1
    #Last Row where data will be examined
    lastRow = 16
    #counter
    rowCounter = rowOffSet
    #flag that indicates Selection was made
    dataAvailable = 0

    while dataAvailable == 0 and rowCounter<lastRow:
        #print(rowCounter)
        #print(colA[rowCounter].value)
        #print(colB[rowCounter].value)
        if dataAvailable == 0 and colB[rowCounter].value == 'X':
            purposeOfTheLetter = colA[rowCounter].value
            dataAvailable = 1
            break
        rowCounter = rowCounter + 1

    # while dataAvailable == 0:
    #     if dataAvailable == 0 and ws['B12'].value == 'X':
    #         purposeOfTheLetter = ws['A12'].value
    #         dataAvailable = 1
    #         break
    #     if dataAvailable == 0 and ws['B13'].value == 'X':
    #         purposeOfTheLetter = ws['A13'].value
    #         dataAvailable = 1
    #         break
    #     if dataAvailable == 0 and ws['B14'].value == 'X':
    #         purposeOfTheLetter = ws['A14'].value
    #         dataAvailable = 1
    #         break
    #     if dataAvailable == 0 and ws['B15'].value == 'X':
    #         purposeOfTheLetter = ws['A15'].value
    #         dataAvailable = 1
    #     if dataAvailable == 0 and ws['B16'].value == 'X':
    #         purposeOfTheLetter = ws['A16'].value
    #         dataAvailable = 1
    #         break
    #end of Collecting purposeOfTheLetter

    ############################
    #Collecting accomplishments
    ############################
    #Initial Row where data will be examined, The offset needs to be set at -1
    rowOffSet = 20 - 1
    #Last Row where data will be examined
    lastRow = 23
    #counter
    rowCounter = rowOffSet
    #flag that indicates Selection was made
    dataAvailable = 0

    while dataAvailable == 0 and rowCounter<lastRow:
        #print(rowCounter)
        #print(colA[rowCounter].value)
        #print(colB[rowCounter].value)
        if dataAvailable == 0 and colB[rowCounter].value == 'X':
            accomplishments = colA[rowCounter].value
            dataAvailable = 1
            break
        rowCounter = rowCounter + 1


    #######################################
    #Collecting Positive Personality Traits
    ######################################

    positivePersonalityTraits = []
    max_number_of_traits = 6
    #Initial Row where data will be examined, The offset needs to be set at -1
    rowOffSet = 27 - 1
    #Last Row where data will be examined
    lastRow = 61
    #counter
    rowCounter = rowOffSet
    #flag that indicates Selection was made
    traitCounter = 0

    while traitCounter <= max_number_of_traits and rowCounter<lastRow:
        #print(rowCounter)
        #print(colA[rowCounter].value)
        #print(colB[rowCounter].value)
        if colB[rowCounter].value == 'X':
            positivePersonalityTraits.append(colA[rowCounter].value)
            traitCounter = traitCounter + 1
        rowCounter = rowCounter + 1

    ############################
    #Collecting Academic Skills
    ############################

    academicSkills = []
    max_number_of_skills = 6
    #Initial Row where data will be examined, The offset needs to be set at -1
    rowOffSet = 66 - 1
    #Last Row where data will be examined
    lastRow = 80
    #counter
    rowCounter = rowOffSet
    #flag that indicates Selection was made
    skillCounter = 0

    while skillCounter <= max_number_of_skills and rowCounter<lastRow:
        if colB[rowCounter].value == 'X':
            academicSkills.append(colA[rowCounter].value)
            skillCounter = skillCounter + 1
        rowCounter = rowCounter + 1




else:
    print("{0} does not appear to be a valid file, choose the right file and retry".format(inputfile))
    sys.exit(2)

print(firstName)
print(lastName)

print(purposeOfTheLetter)

print(accomplishments)

print(positivePersonalityTraits)

print(academicSkills)
