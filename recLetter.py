import Listofthings

from Listofthings import PronoumsList,PositivePersonalityTraits,AcademicSkills, Phrase1, Phrase2, Phrase3, Phrase4, Phrase5, LinkingWords
import openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook
import sys, getopt
import os
import random
import datetime
from datetime import date


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
    pronoumLocalList = PronoumsList.get(pronoum)
    #he
    subjective = pronoumLocalList[0]
    #him
    objective = pronoumLocalList[1]
    #his
    possessive = pronoumLocalList[2]

    classAttended = ws['B6'].value

    highSchoolYearAttended = ws['B8'].value


    ###########################
    #Time teacher know students
    ###########################

    schoolYearAttended = ws['B7'].value
    yearsAttended = schoolYearAttended.split('/')
    start_date = datetime.datetime(int(yearsAttended[0]), 8, 15)
    num_months = (date.today().year - start_date.year) * 12 + (date.today().month - start_date.month)

    targettedInstitution = ws['B9'].value



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


    ############################
    #Collecting accomplishments
    ############################
    accomplishments = 0
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




_space = " "

################
#1st Paragraph
################

print("\n\rTo whom it may concern, \n\r\n\r")

                #"It is with pleasure that I recommend "                        "I was fortunate to have" "him"
print (random.choice(Phrase1) + firstName + _space + lastName + " to the " + targettedInstitution + _space + purposeOfTheLetter + ". " +
random.choice(Phrase2) + _space + objective.lower() + _space + "in my classroom. "+ firstName + " was a "
+ highSchoolYearAttended + " in my " + classAttended + " class in " + schoolYearAttended + ". ", end="", flush=True)

if(num_months < 12) : print( "Although I have only taught " + firstName + _space + "for " + str(int(num_months)) + " months, I can already see ", end="", flush=True)
if(num_months > 12 and num_months < 24) : print("I have known " + firstName + _space + "for over an year, and " + subjective.lower() + " made an impression in me due to ", end="", flush=True)
if(num_months > 24) : print("I have known " + firstName + _space + "for more than " + str(int(num_months/12)) + " years, and " + subjective.lower() + " made an impression in me due to ", end="", flush=True)

#print(possessive.lower() + _space + positivePersonalityTraits[0] + _space + "and also " + positivePersonalityTraits[1] + " personality.", end="", flush=True)
print(possessive.lower() + _space + random.choice(academicSkills) + _space + "and also " + random.choice(positivePersonalityTraits) + " personality. ", end="", flush=True)

#"I want to illustrate a little more about ",         #Him/Her
print(random.choice(Phrase3) +  objective.lower() + " in this letter, and why " + subjective.lower() +  " deserves to be considered in your instituition.")

print("\n\r\n\r")

################
#2nd Paragraph
################

ACSkills1 = random.choice(academicSkills)
academicSkills.remove(ACSkills1)
ACSkills2 = random.choice(academicSkills)
academicSkills.remove(ACSkills2)
ACSkills3 = random.choice(academicSkills)
academicSkills.remove(ACSkills3)
ACSkills4 = random.choice(academicSkills)
academicSkills.remove(ACSkills4)
ACSkills5 = random.choice(academicSkills)
academicSkills.remove(ACSkills5)
ACSkills6 = random.choice(academicSkills)
academicSkills.remove(ACSkills6)



print("While in class I have observed some remarkable academic skills. " + firstName + _space + AcademicSkills[ACSkills1] + ". " + subjective  + " also " + AcademicSkills[ACSkills2]
+ random.choice(Phrase5) + subjective.lower() + _space + AcademicSkills[ACSkills3]+ random.choice(LinkingWords) + subjective.lower() + _space  + AcademicSkills[ACSkills4] + ".")

print("\n\r\n\r")
#print(firstName)
#print(lastName)

#print(purposeOfTheLetter)

#print(accomplishments)

#print(positivePersonalityTraits)

#print(academicSkills)
