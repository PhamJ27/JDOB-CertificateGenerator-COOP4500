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
from docx.oxml.ns import qn
import time

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

Template = 'AwardTemplate2019-2020.docx'
SpreadSheet = 'Processed_Report1-Revised.csv'

def all_paragraphs(doc):
    for p in doc.element.body.iter(qn('w:p')):
        yield Paragraph(p, doc)
    # gives all the paragraphs within a doc, pausing and sending after every one

def deTERMine(month):
    if month.isdigit():
        if      ( int(month) == 9 or (int(month) == 10) or (int(month) == 11) or int(month) == 12 ):
            return 1
        elif    ( (int(month) == 1) or (int(month) == 2) or int(month) == 3 ):
            return 2
        elif    ( int(month) == 4 or (int(month) == 5) or (int(month) == 6) ):
            return 3
        else:
            return 4
    elif isinstance(month, str):
        if month in monthDict.values():
            if      ( month in (monthDict[9], monthDict[10], monthDict[11], monthDict[12] ) ):
                return 1
            elif    ( month in (monthDict[1], monthDict[2], monthDict[3] ) ):
                return 2
            elif    ( month in (monthDict[4], monthDict[5], monthDict[6] ) ):
                return 3
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


def schoolYearCalc(year, month):
    if month.isdigit() and month.isdigit():
        month_num = int(month)
    elif isinstance(month, int):
        month_num = month
    elif isinstance(month, str):
        month_num = next(k for k, v in monthDict.items() if v == month)
    else:
        print("ERROR: MONTH IS OUT OF BOUNDS AND NOT A VALID MONTH.")
        return -1
    year = int(year)

    if 9 <= month_num <= 12: # SEPT to DEC
        return f"{year} - {year + 1}"
    elif 1 <= month_num <= 8: # JAN TO AUG
        return f"{year - 1} - {year}"
    else:
        print("ERROR: MONTH IS OUT OF BOUNDS AND NOT A VALID MONTH.")
        return -1

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
#given a paragraph, as long it's true (so replace all placeholders).
#given a list of bodies of text, join all of them to full.
#the start is when the placeholder is first found.
#the end is start + length of the placeholder. (so the whole placeholder)
# for every body of text, if a placeholder is directly before or after, ignore it.
#otherwise if the body of text overlaps,
# the start is start or run_start (whichever is bigger) - run_start
# and our end is end or run_end (whichever is smaller) - run_start

#if it hasn't been placed, replace it and mark it


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
            print(SFN + " - " + Certificate)
            assign(SFN, Certificate, Month, Years, LastName)
    #reads the sheet for names.
    # return

def assign(StudentFullName, Certificate, Month, Year, LastName):
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

    output_folder = Path(f"{month:02d}" + "-" + Years)
    output_folder.mkdir(parents=True, exist_ok=True)

    FirstName = StudentFullName.split()[0]
    filename = (LastName + FirstName + "-" + f"{month:02d}" + "-" + Years) + ".docx"
    fullPath = output_folder / filename

    doc.save(fullPath)
    #writes the name into the doc
    return


# Press the green button in the gutter to run the script.
if __name__ == '__main__':

    # fileCheck()

    # year = 2026
    # month = 9
    YearOfGraduation = 2030
    #manually inputted year and month

    year = datetime.datetime.now().year
    month = datetime.datetime.now().month
    #year and month grabbed from NOW

    SchoolYear = str(year) + " - " + str(year + 1)
    Years = str(year)
    # monthCorrected = monthList[month]
    monthCorrected = monthDict.get(month)       #retrieve

    # print(Years)
    # print(monthCorrected)

    readsheet(SpreadSheet, monthCorrected, Years)
    # fileCheck()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
