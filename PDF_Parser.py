from os import listdir
from os.path import isfile, join, isdir
import xlsxwriter
import re
import PyPDF2
import datetime
import os
import time


# I'm well aware this program can be glitchy.  It's not very flexible at all.  It was created soley to parse transmittal files and thus is contingent on having transmittal files follow a certain pattern.  I tried to account for as much deviation as possible, but I may not have captured it all.


# Global variable of the working compiled scraped data from the pdf documents
compiled = {}

# Flag to mark the end of PDF parsing
endFlag = "FO=Full"

# REGEX patterns for dates, transmittal names, dates, and parsing glitches
transmittalPattern = r"([a-zA-Z]*[0-9]{4}).pdf"
transmittalPattern2 = r"T-IONE"
datePattern = r"([0-9]{1,2})/([0-9]{1,2})/([0-9]{2,4})"
datePattern2 = r"([0-9]{1,2})-([a-zA-Z]{3})-([0-9]{2,4})"
dateMatch1 = r"([0-9]{1,2})/([0-9]{1,2})/([0-9]{4})"
dateMatch2 = r"([0-9]{1,2})/([0-9]{1,2})/([0-9]{2})"
dateMatch3 = r"([0-9]{1,2})-([a-zA-Z]{3})-([0-9]{4})"
dateMatch4 = r"([0-9]{1,2})-([a-zA-Z]{3})-([0-9]{2})"
glitchPattern = r"(-)\s"
glitchPattern2 = r"\s(\))"

def main():

    srcPath = ""

    # Request input for path to scrape
    exists = False
    srcPath = input("SET SCRAPING SOURCE DIRECTORY: ")
    if os.path.exists(srcPath):
        exists = True
    else:
        exists = False

    while(not exists):
        srcPath = input("directory doesn't exist. Please enter source directory: ")
        if os.path.exists(srcPath):
            exists = True
        else:
            exists = False

    print(srcPath)
    exists = False

    destPath = input("SET OUTPUT DIRECTORY: ")
    if os.path.exists(destPath):
        exists = True
    else:
        exists = False


    while(not exists):
        destPath = input("directory doesn't exist. Please enter destination directory: ")
        if os.path.exists(destPath):
            exists = True
        else:
            exists = False


    # Set row on excel worksheet to 1 (second row. first row is for headers)
    row = 1

    # run fileScrape method on the set path
    fileScrape(srcPath)

    #
    book = xlsxwriter.Workbook(destPath + "\\PDF_Dump.xlsx")
    link_format = book.add_format({'color':'blue'})
    sh = book.add_worksheet()
    # Setting the header
    sh.write(0,0,"Supplier")
    sh.write(0,1,"Transmittal Number")
    sh.write(0,2,"Transmittal Name")
    sh.write(0,3,"Date")
    sh.write(0,4,"Text Body")

    for transmittalKey in compiled.keys():
        date = compiled[transmittalKey]['date']
        tempStr = ' '.join(compiled[transmittalKey]['body'])
        tempStr = ' '.join(tempStr.split())
        tempStr = tempStr.replace("- ","-")
        tempStr = tempStr.replace(" )",")")
        print("printing " + transmittalKey + " to worksheet...")
        sh.write(row,0,compiled[transmittalKey]["supplier"])
        sh.write(row,1,compiled[transmittalKey]["transNum"])
        sh.write_url(row,2,compiled[transmittalKey]['link'],link_format,transmittalKey)

        if (re.match(dateMatch1,date) != None) & (validateDate(date,'%m/%d/%Y')):
            date_time = datetime.datetime.strptime(date,'%m/%d/%Y')
        elif (re.match(dateMatch2,date) != None) & (validateDate(date,'%m/%d/%y')):
            date_time = datetime.datetime.strptime(date,'%m/%d/%y')
        elif (re.match(dateMatch3,date) != None) & (validateDate(date,'%d-%b-%Y')):
            date_time = datetime.datetime.strptime(date,'%d-%b-%Y')
        elif (re.match(dateMatch4,date) != None) & (validateDate(date,'%d-%b-%y')):
            date_time = datetime.datetime.strptime(date,'%d-%b-%y')
        else:
            date = (time.strftime('%m/%d/%Y',time.gmtime(compiled[transmittalKey]['cdate'])))
            date_time = datetime.datetime.strptime(date,'%m/%d/%Y')

        if date != "":
            date_format = book.add_format({'num_format':'mm/dd/yyyy'})
            sh.write_datetime(row,3,date_time,date_format)
        else:
            sh.write(row,3,"")
        # sh.write(count,3,compiled[transmittalKey]["date"])

        sh.write(row,4,tempStr)
        row += 1


    book.close()
    print("Completed")
    x = input("press any key to continue")

# validate the date
def validateDate(dateString, dateFormatTest):
    try:
        datetime.datetime.strptime(dateString,dateFormatTest)
        return True
    except ValueError:
        return False

# recursively dig through tree of folders and scrape for PDF files.  PDF files are sent to the pdfDataCollect method.
def fileScrape(rootPath):
    for f in listdir(rootPath):
        if(isfile(join(rootPath,f)) & f.endswith("pdf") & ((re.match(transmittalPattern,f) != None) or (re.match(transmittalPattern2,f) != None))):
            print(f)
            pdfDataCollect(join(rootPath,f),f)
        elif(isdir(join(rootPath,f))):
            print("dir " + f)
            fileScrape(join(rootPath,f))


# Parse through PDF files.  Pull the date and text body from the pdf files.  Pull the transmittal number, supplier name and file link from the file path.
def pdfDataCollect(pdfPath,f):
    dateTrigger = 0
    wordArr = []
    newWordArr = []
    date = ''
    fileLink = pdfPath

    pdfReader = PyPDF2.PdfFileReader(open(pdfPath,'rb'))
    for x in range(0,pdfReader.numPages):
        pageObj = pdfReader.getPage(x)
        wordArr += pageObj.extractText().splitlines()

# Parsing through the pdf
    for word in wordArr:
        if (re.search(datePattern,word) != None) & (dateTrigger == 0):
            date = word[re.search(datePattern,word).span()[0]:re.search(datePattern,word).span()[1]]
            dateTrigger = 1
        if (re.search(datePattern2,word) != None) & (dateTrigger == 0):
            date = word[re.search(datePattern2,word).span()[0]:re.search(datePattern2,word).span()[1]]
            dateTrigger = 1
        if (re.search(endFlag,word) != None):
            break
        else:
            newWordArr.append(word)

# pull the date the file was created.
    cdate = os.path.getctime(pdfPath)

# store data in dictionary in a dictionary.  The data is structure as compiled -> transmittal name (i.e. 'WEIR0032') -> transmittal number, supplier name, pdf text, date, created date, file path
    nameKey = f[:-4]

    compiled[nameKey] = {
        "transNum":f[-8:-4],
        "supplier":f[0:-8],
        "body":newWordArr,
        "date":date,
        "cdate":cdate,
        "link":fileLink
    }

if __name__ == "__main__":
    main()