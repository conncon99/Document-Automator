import xlwings as xw
from docxtpl import DocxTemplate
from time import perf_counter, time
import pandas as pd
import os
import projectManager as pM


def setFieldsOnSheet(neededFields: list, loanPackName: str, sheet: xw.Book):
    loanSheet = xw.Book(pM.workBookName).sheets[loanPackName]
    translatorDict = loanSheet.range(pM.translationTableCell).options(dict, expand='table', numbers=str).value
    pM.checkInvalidKeys(translatorDict, sheet)
    inverted_dict = dict(map(reversed, translatorDict.items()))
    neededFields = [inverted_dict[field] for field in neededFields]
    
    for i in range(pM.maxLength):
        loanSheet.range('C' + str(4 + i)).value = neededFields[i] if i < len(neededFields) else "filler"

def emptyCRow(loanPackName: str): 
    loanSheet = xw.Book(pM.workBookName).sheets[loanPackName]
    for i in range(pM.maxLength):
        loanSheet.range('C' + str(4 + i)).value = "filler"
        
def createOrUpdateCSV(templatePath: str, unsavedDocuments: list, loanPackName:str, exists: bool, pathCSV: str, sheet: xw.Book):
    if not os.path.exists(f"{templatePath}/{loanPackName}"):
        raise pM.InvalidPackName(sheet)
    df = pd.read_csv(pathCSV) if exists else pd.DataFrame()
    loanPath = f"{templatePath}/{loanPackName}/"
    allFields = set()
    for fileName in unsavedDocuments:
        filePath = f"{loanPath}{fileName}"
        if not os.path.exists(filePath):
            pM.hasMatchingFileName(fileName, loanPath)
        sheet[pM.workingFileCell].value = fileName
        foundFields = set()
        doc = DocxTemplate(filePath)
        foundFields.update(doc.undeclared_template_variables)
        adjustedFields = sorted(foundFields)
        for i in range(pM.maxLength - len(foundFields)):
            adjustedFields.append("filler")
        df[fileName] = pd.Series(adjustedFields)
        allFields = allFields.union(foundFields)
    df.to_csv(pathCSV, index=False)
    return allFields
        

        
def getFields(baseDir: str, loanPackName: str, selectedDocuments: list, sheet: xw.Book):
    neededFields = set()
    templatePath = baseDir + "/Templates"
    selectedDocuments = [fileName + ".docx" for fileName in selectedDocuments]
    unsavedDocuments = selectedDocuments[:]
    pathCSV = f"{templatePath}/{loanPackName}.csv"
    
    if os.path.exists(pathCSV) and os.path.getmtime(pathCSV) < time() - 3600:
        df = pd.read_csv(pathCSV)
        for fileName in selectedDocuments:
            #Checks if the file is in the csv already and whether it has been modified in the last hour - the template files should not need to be modified but just in case
            if fileName in df.columns:
                sheet[pM.workingFileCell].value = fileName
                neededFields.update(df[fileName].to_list())
                unsavedDocuments.remove(fileName)  
        neededFields.update(createOrUpdateCSV(templatePath, unsavedDocuments, loanPackName, True, pathCSV, sheet))
    else:
        neededFields.update(createOrUpdateCSV(templatePath, unsavedDocuments, loanPackName, False, pathCSV, sheet))
    #Check CSV matches Translator
    pM.checkTranslationMatchesCSV(pathCSV, xw.Book(pM.workBookName).sheets[loanPackName], sheet)
    if "filler" in neededFields:
        neededFields.remove("filler")
    return sorted(neededFields)
            
def main():
    startTime = perf_counter()
    wb = xw.Book.caller()
    sheet, baseDir, loanPackName = pM.setupSpreadsheet(wb)
    emptyCRow(loanPackName)
    selectedDocuments = pM.getSelectedDocuments(sheet.range(pM.selectedDocumentsTableCell).options(dict, expand='table').value)
    if selectedDocuments != []:
        neededFields = getFields(baseDir, loanPackName, selectedDocuments, sheet)
        setFieldsOnSheet(neededFields, loanPackName, sheet)
    pM.onComplete(startTime, perf_counter(), sheet)

if __name__ == "__main__":
    xw.Book(pM.workBookName).set_mock_caller()
    main()
    print("Done!")
