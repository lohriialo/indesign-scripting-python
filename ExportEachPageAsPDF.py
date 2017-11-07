"""
Exports each page of an InDesign document as a separate PDF to
a specified folder using the current PDF export settings.
"""

import win32com.client
import os

app = win32com.client.Dispatch('InDesign.Application.CC.2017')

myFile = r'C:\ServerTestFiles\TestDocument.indd'
myDocument = app.Open(myFile)

idPDFType = 1952403524
if app.Documents.Count is not 0:
    directory = os.path.dirname(myFile)
    docBaseName = myDocument.Name
    for x in range(0, myDocument.Pages.Count):
        myPageName = myDocument.Pages.Item(x + 1).Name
        # We want to export only one page at the time
        app.PDFExportPreferences.PageRange = myPageName
        # strip last 5 char(.indd) from docBaseName
        myFilePath = directory + "\\" + docBaseName[:-5] + "_" + myPageName + ".pdf"
        myDocument.Export(idPDFType, myFilePath)

myDocument.Close()
