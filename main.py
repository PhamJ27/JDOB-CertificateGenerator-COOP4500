# This is a sample Python script.
import datetime
# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.


from multiprocessing import Pool
import csv
from docx import Document
import os
from pathlib import Path
from docx.shared import Pt

from docx.text.paragraph import Paragraph
from docxcompose.composer import Composer
from docx.oxml.ns import qn
from docx2pdf import convert
import time


Template    = 'AwardTemplate2019-2020.docx'
SpreadSheet = 'Processed_Report1-Revised.csv'

# monthList = [
    #                   "January", "February", "March",
    #                    "April", "May", "June",
    #                    "July", "August", "September",
    #                    "October", "November", "December"
    #             ]

monthDict = {
    1 : "January",  2 : "February",     3 : "March",
    4 : "April",    5 : "May",          6 : "June",
    7 : "July",     8 : "August",       9 : "September",
    10 : "October", 11 :"November",     12 : "December"
}

year = datetime.datetime.now().year
month = datetime.datetime.now().month

studentcount    = 0
HighH           = 0
HRolls          = 0
HMention        = 0

needInfo = False

def statcounter():
    global studentcount
    global HighH
    global HRolls
    global HMention
    global SpreadSheet

    with open(SpreadSheet, encoding='utf-8-sig') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            Certificate = row['Certificate']
            studentcount = + studentcount + 1
            if  (Certificate == "High Honor's"):
                HighH += 1
            elif (Certificate == "Honorable Mention"):
                HMention += 1
            elif (Certificate == "Honor Roll"):
                HRolls += 1
    #read the spreadsheet, read each row. count each student. then count each of the Certificate type.

def yearCalc(yearOfGraduation):
    global year
    yearOfGraduation = int(yearOfGraduation)
    yearsLeft = yearOfGraduation - year
    match (yearsLeft):
        case 5:
            return 7
        case 4:
            return 8
        case 3:
            return 9
        case 2:
            return 10
        case 1:
            return 11
        case 0:
            return 12
        case _:
            print("Years left out of bounds.")
            return -1
    #calculate the grade using year of graduation and the year.
    # year is calculated previously as the current year, but can be changed

def all_paragraphs(doc):
    for p in doc.element.body.iter(qn('w:p')):
        yield Paragraph(p, doc)
    # gives all the paragraphs within a doc, pausing and sending after every one

def deTERMine(month):
    if isinstance(month, int):
        if      ( 9 <= int(month) <= 11 ):                      #SEPT to NOV
            return 1
        elif    ( 1 <= int(month) <= 2 or int(month) == 12 ):   #DEC to JAN
            return 2
        elif    ( 3 <= int(month) <= 4 ):                       #MAR to APR
            return 3
        elif    ( 5 <= int(month) <= 6 ):                       #MAY to JUN
            return 4
        else:
            return -1                                           #not valid - they are off by a month for leniency
    elif isinstance(month, str):
        if month in monthDict.values():
            if      ( month in (monthDict[9], monthDict[10], monthDict[11]) ):  #SEPT to NOV
                return 1
            elif    ( month in (monthDict[12], monthDict[1], monthDict[2] ) ):  #DEC to JAN
                return 2
            elif    ( month in (monthDict[3], monthDict[4] ) ):                 #MAR to APR
                return 3
            elif    (month in (monthDict[5], monthDict[6]) ):                   #MAY to JUN
                return 4
            else:
                return -1
        else :
            # print(month)
            print("ERROR: MONTH IS A STRING, BUT NOT A VALID MONTH.")
            return -1
    print("ERROR: MONTH IS NOT A STRING NOR A DIGIT")
    return -1

def deTERMineSUFFIX(month):
    if deTERMine(month):
        if      ( deTERMine(month) == 1 ):
            return "st"
        elif    ( deTERMine(month) == 2 ):
            return "nd"
        elif    ( deTERMine(month) == 3):
            return "rd"
        elif    ( deTERMine(month) == -1 ):
            print("ERROR: ")
        else:   # or (int(month) == 5) or (int(month) == 6) #tbh works with everything else
            return "th"
    else:
        print("ERROR: NOT A MONTH TO APPLY SUFFIX")


def schoolYearCalc(year, month):
    if month.isdigit():
        month_num = int(month)
    elif isinstance(month, int):
        month_num = month
    elif isinstance(month, str):
        month_num = next(k for k, v in monthDict.items() if v == month)
    else:
        print("ERROR: MONTH IS OUT OF BOUNDS AND NOT A VALID MONTH.")
        return -1
    year = int(year)
    if 9 <= month_num <= 12:    # SEPT to DEC
        return f"{year} - {year + 1}"
    elif 1 <= month_num <= 8:   # JAN TO AUG
        return f"{year - 1} - {year}"
    else:
        print("ERROR: MONTH IS OUT OF BOUNDS AND NOT A VALID MONTH.")
        return -1
    # taking in the month as any data type, the current year range is based on the month you're in
    # from SEPT to DEC, the schoolyear is current & current + 1
    # from JAN to AUG to school year is current - 1 & current

def replace_in_paragraph(paragraph, old, new, font_name, size, bold, superscript):
    while True:
        runs = paragraph.runs
        full = "".join(r.text for r in runs)
        start = full.find(old)
        if start == -1:
            return
        end = start + len(old)

        pos = 0
        placed = False
        for run in runs:
            run_start, run_end = pos, pos + len(run.text)
            pos = run_end
            if run_end <= start or run_start >= end:
                continue

            s = max(start, run_start) - run_start
            e = min(end, run_end) - run_start
            if not placed:
                run.text = run.text[:s] + new + run.text[e:]
                run.font.name = font_name
                run.font.size = Pt(size)
                run.font.bold = bold
                run.font.superscript = superscript  # new line
                placed = True
            else:
                run.text = run.text[:s] + run.text[e:]
# paragraphs are made of runs
# a run is continuous piece of text with the same format
#  use the full text to find the start. the end is end of old
#  for every run its start and end is 0 and end of the run's text.
# replace that bulk, then mark it as true.
# now the run = rest of the start - bulk - rest of the end


def fileCheck():
    print(os.getcwd())
    print(os.path.exists(SpreadSheet))
    print(os.path.exists(Template))


def readsheet(sheet, Month, Years):
    with open(sheet, encoding='utf-8-sig') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            SFN = row['StudentFullName']
            Certificate = row['Certificate']
            LastName = SFN.split()[-1]
            YearOfGraduation = row['Year']
            print(SFN + " - " + Certificate)
            assign(SFN, Certificate, Month, Years, LastName, YearOfGraduation)
    # reads the sheet for names.
    # open the file, read it.
    # for every row, extract the full name, certificate, last name, and Year of Graduation.
    # print the current student.
    # then assign it.

def assign(StudentFullName, Certificate, Month, Year, LastName, YearOfGraduation):
    # string, string, string, string, string
    doc = Document(Template)

    termNumber  = str(deTERMine(Month))
    Suffix      = deTERMineSUFFIX(Month)
    SchoolYear  = schoolYearCalc(Year, Month)

    # old_text             Replacement, Font, Size, Boldness, Superscript
    replacements = {
        "{{FULLNAME}}":     (StudentFullName,   "Monotype Corsiva", 47, True, False),
        "{{CERTIFICATE}}":  (Certificate,       "Monotype Corsiva", 47, True, False),
        "{{MONTH}}":        (Month,             "Monotype Corsiva", 22, True, False),
        "{{YEAR}}":         (Year,              "Monotype Corsiva", 22, True, False),
        "{{TERMNO}}":       (termNumber,        "Monotype Corsiva", 18, False, False),
        "{{SUFFIX}}":       (Suffix,            "Monotype Corsiva", 18, False, True),
        "{{SCHOOLYEAR}}":   (SchoolYear,        "Monotype Corsiva", 18, False, False)
    }

    for paragraph in all_paragraphs(doc):
        for old_text, (new_text, font_name, size, bold, superscript) in replacements.items():
            replace_in_paragraph(paragraph, old_text, new_text, font_name, size, bold, superscript)

    # for every paragraph and
    # for every old piece of text and element to replace in the replacement list (shown as replaced, (replacement, items)
    # replace it, using the text, font, size, boldness, and superscript.

    output_folder = Path("Grade" + f"{yearCalc(YearOfGraduation):02d}" + " - " + Years + " - " + "Term" + termNumber)
    output_folder.mkdir(parents=True, exist_ok=True)    #creates parents and does nothing if it already exists.

    FirstName = StudentFullName.split()[0]
    filename = ("Grade" + f"{yearCalc(YearOfGraduation):02d}" + " - " + Certificate + " - " + LastName + FirstName) + ".docx"
    fullPath = output_folder / filename

    # OutputFolder = RelativePathToProgram / Grade - GradeNumber - YearRange - Term - TermNumber
    # StudentName = Grade + GradeNo. + CertificateType + Lastname + Firstname

    doc.save(fullPath)
    #writes the name into the doc

    # return doc
    # returns the newly created document


# THE NEXT FUNCTIONS ARE MERGED VERSIONS OF THE PREVIOUS ONES. THEY WORK IN TANDEM TO CREATE ONE SINGLE FILE.

def readsheetmerge(sheet, Month, Years):
    docs = []
    with open(sheet, encoding='utf-8-sig') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            SFN = row['StudentFullName']
            Certificate = row['Certificate']
            LastName = SFN.split()[-1]
            YearOfGraduation = row['Year']
            print(SFN + " - " + Certificate)
            student_doc = assignMerge(SFN, Certificate, Month, Years, LastName, YearOfGraduation)
            docs.append(student_doc)
    return docs
    #reads the sheet for names. read readsheetmerge() for the rest of the comments
    # return

def assignMerge(StudentFullName, Certificate, Month, Year, LastName, YearOfGraduation):
    # string, string, string, string, string
    doc = Document(Template)

    termNumber  = str(deTERMine(Month))
    Suffix      = deTERMineSUFFIX(Month)
    SchoolYear  = schoolYearCalc(Year, Month)

    # toReplace             Replacement, Font, Size, Boldness, Superscript
    replacements = {
        "{{FULLNAME}}":     (StudentFullName,   "Monotype Corsiva", 47, True, False),
        "{{CERTIFICATE}}":  (Certificate,       "Monotype Corsiva", 47, True, False),
        "{{MONTH}}":        (Month,             "Monotype Corsiva", 22, True, False),
        "{{YEAR}}":         (Year,              "Monotype Corsiva", 22, True, False),
        "{{TERMNO}}":       (termNumber,        "Monotype Corsiva", 18, False, False),
        "{{SUFFIX}}":       (Suffix,            "Monotype Corsiva", 18, False, True),
        "{{SCHOOLYEAR}}":   (SchoolYear,        "Monotype Corsiva", 18, False, False)
    }

    for paragraph in all_paragraphs(doc):
        for old_text, (new_text, font_name, size, bold, superscript) in replacements.items():
            replace_in_paragraph(paragraph, old_text, new_text, font_name, size, bold, superscript)
    return doc
    # returns the newly created document, read assign() for the rest of the comments

if __name__ == '__main__':

    # fileCheck()

    # year = 2025
    # month = 5
    # YearOfGraduation = 2030
    #manually inputted year and month

    # global month
    # global year
    #year and month grabbed from NOW

    SchoolYear = str(year) + " - " + str(year + 1)
    Years = str(year)
    # monthCorrected = monthList[month]
    monthCorrected = monthDict.get(month)       #retrieve

    # print(Years)
    # print(monthCorrected)

    # readsheet(SpreadSheet, monthCorrected, Years)
    # fileCheck()

    all_docs = readsheetmerge(SpreadSheet, monthCorrected, Years)

    if all_docs:
        master = Composer(all_docs[0])
        for doc in all_docs[1:]:
            master.doc.add_page_break()  # ensure each student starts on a new page
            master.append(doc)
        with open(SpreadSheet, encoding='utf-8-sig') as csvfile:
            reader = csv.DictReader(csvfile)
            first_row = next(reader)
            YearOfGraduation = first_row['Year']    #just grab the first row's year
        termNumber = str(deTERMine(month))

        output_folder = Path("output")
        output_folder.mkdir(parents=True, exist_ok=True)

        filename = ("Grade" + f"{yearCalc(YearOfGraduation):02d}" + " - " + Years + " - " + "Term" + termNumber)
        # Grade - GradeLevels( 08, 09, 10, 11, 12) - Year Range - Term - TermNumber (1, 2, 3, 4)
        master.save(output_folder / (filename + ".docx"))
        convert((output_folder / (filename + ".docx") ), (output_folder / (filename + ".pdf")  ))
    # until it's done reading all_docs
    # there's a list of docs called masterdocs using composer - (from docxcompose)
    # this doc skips to the 2nd student, adds a page break, then appends the doc.
    # the rest is naming, saving, and converting the file.
    # merged version

    print("Certificates printed from: " + SpreadSheet + "are complete. They are stored in the output folder.")
    if needInfo:
        statcounter()
        print(f"Number of Students: {studentcount}")
        print(f"Number of High Honors {HighH}")
        print(f"Number of Honor Rolls {HRolls}")
        print(f"Number of Honorable Mentions: {HMention}")

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
