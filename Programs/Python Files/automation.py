import xlwings as xw
from docxtpl import DocxTemplate
from time import perf_counter
from multiprocessing import Pool, freeze_support
from functools import partial
import projectManager as pM
import os
import jinja2

def writeToDocuments(selectedDocuments: list, loanPackName: str, baseDir: str):
    #Set up multiprocessing of several documents at once
    with Pool(4) as p:
        p.map(partial(writeDocument, loanPackName, baseDir),selectedDocuments)
        
def writeDocument(loanPackName: str, baseDir: str, fileName: str):
    #xlWings module is bugged so have to redefine sheet inside multiprocessing
    sheet = xw.Book(pM.workBookName).sheets["Master"]
    loanPackPath = f"{baseDir}/Templates/{loanPackName}"
    
    if not os.path.exists(loanPackPath):
        raise pM.InvalidPackName(sheet)
    loanSheet = xw.Book(pM.workBookName).sheets[loanPackName]
    fileName += '.docx'
    
    if not os.path.exists(loanPackPath +"/" + fileName):
        pM.hasMatchingFileName(fileName, loanPackPath + "/")
    
    doc = DocxTemplate(f"{loanPackPath}/{fileName}")
    sheet.range(pM.workingFileCell).value = fileName
    #Take inputs from sheet and convert them to be usable by docxtpl
    inputDict = sheet.range(pM.fieldTableCell).options(dict, expand='table', numbers=str).value
    translatorDict = loanSheet.range(pM.translationTableCell).options(dict, expand='table', numbers=str).value
    pM.checkInvalidKeys(translatorDict, sheet)
    translatorDict = dict((translatorDict[key],value) for (key,value) in inputDict.items())
    #Create new document - most of the performance issues comes from doc.render, this does not appear to be significantly affected by the dictionary size, which makes optimization hard
    doc.render(translatorDict, autoescape=True)
    #Save the new document to the output folder
    doc.save(f"{baseDir}/Output/{fileName}")
        
def main():
    #Accesses excel sheet
    startTime = perf_counter()   
    wb = xw.Book.caller()
    
    sheet, baseDir, loanPackName = pM.setupSpreadsheet(wb)    
    selectedDocuments = pM.getSelectedDocuments(sheet.range(pM.selectedDocumentsTableCell).options(dict, expand='table').value)
    writeToDocuments(selectedDocuments, loanPackName, baseDir)
    
    pM.onComplete(startTime, perf_counter(), sheet)

if __name__ == "__main__":
    #compatability with pyinstaller and multiprocessing
    freeze_support()
    xw.Book(pM.workBookName).set_mock_caller()
    main()
    print("Done!")
