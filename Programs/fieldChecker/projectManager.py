import xlwings as xw
import os
import pandas as pd

#Shared Fields
#Master Sheet
programStatusCell = 'G8'
workingFileCell = 'G9'
timeTakenCell = 'G10'
errorMessageCell = 'E11'
loanPackCell = 'C2'
selectedDocumentsTableCell = 'L5'
pathCell = 'A1'
fieldTableCell = 'B4'

#Loan Pack Sheet
translationTableCell = 'K3'

maxLength = 50

workBookName = "automation.xlsm"

#Util Methods

def setupSpreadsheet(wb: xw.Book):
    wb.display_alerts = False
    sheet = wb.sheets["Master"]
    sheet[errorMessageCell].value = ""
    baseDir = sheet.range(pathCell).value
    loanPackName = sheet.range(loanPackCell).value
    return sheet, baseDir, loanPackName

def getSelectedDocuments(documentDict: dict):
    selectedDocuments = []
    for key in documentDict.keys():
        if documentDict.get(key) == "Yes" and key != 'filler':
            selectedDocuments.append(key)
    return selectedDocuments

def onComplete(startTime, stopTime, sheet: xw.Book):
    sheet.range(programStatusCell).value = 'Complete!'
    sheet.range(workingFileCell).value = ""
    sheet.range(timeTakenCell).value = stopTime - startTime
    
#Error Checking Methods

def hasMatchingFileName(fileName: str, path: str):
    invalidExtensions = ['pdf','doc', 'html', 'htm', 'odt', 'xls', 'xlsx', 'ods', 'txt', 'ppt', 'pptx']
    for extension in invalidExtensions:
        wrongExtensionFile = fileName.replace('docx', extension)
        if os.path.exists(path + wrongExtensionFile):
            raise InvalidFileType(wrongExtensionFile)
    raise InvalidFileName(fileName)

def checkInvalidKeys(translatorDict: dict, sheet: xw.Book):
    for key in translatorDict.keys():
        if translatorDict[key] == "" or translatorDict[key] == "filler" and key != "filler":
            raise TranslatorError(key, sheet)
        
def checkTranslationMatchesCSV(pathCSV: str, loanPackSheet: xw.Book , sheet: xw.Book):
    df = pd.read_csv(pathCSV)
    csvSet = set(pd.unique(df.values.ravel('K')))
    csvSet.add('MachinePlaceHolder')
    translatorSet = set(loanPackSheet.range(translationTableCell).options(dict, expand='table', numbers=str).value.values())
    
    if not translatorSet.issuperset(csvSet):
        raise InvalidField(str(csvSet.difference(translatorSet)), sheet)
    
        
#Exception Classes
    
class ExcelException(Exception):
    def __init__(self, message: str, sheet = xw.Book(workBookName).sheets["Master"]):
        sheet[errorMessageCell].value = message
        super().__init__(message)
    
class InvalidFileName(ExcelException):
    def __init__(self, fileName: str):
        super().__init__(f"{fileName} cannot be found")
        
class TranslatorError(ExcelException):
    def __init__(self, key, sheet):
        super().__init__(f"Field {key} in {sheet[loanPackCell].value} translation table is blank/filler")
        
class InvalidField(ExcelException):
    def __init__(self, fields: set, sheet: xw.Book):
        super().__init__(f"{fields} could not be found in the {sheet[loanPackCell].value} translation table", sheet)
        
class InvalidPackName(ExcelException):
    def __init__(self, sheet: xw.Book):
        super().__init__(f"The {sheet[loanPackCell].value} sheet does not have a matching path in the Templates folder", sheet)
        
class InvalidFileType(ExcelException):
    def __init__(self, fileName: str):
        super().__init__(f"The file {fileName} must have the extension .docx")