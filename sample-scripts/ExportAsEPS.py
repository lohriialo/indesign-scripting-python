"""
Use the DocumentPreferences object to change the
dimensions and orientation of the document
"""

import win32com.client
import os

app = win32com.client.Dispatch('InDesign.Application.CC.2017')

myFile = r'C:\ServerTestFiles\TestDocument.indd'
myDocument = app.Open(myFile)

idEPSType = 1952400720
if app.Documents.Count is not 0:
    myExportedEPSFile = r'C:\ServerTestFiles\TestDocument.eps'
    directory = os.path.dirname(myExportedEPSFile)
    if not os.path.exists(directory):
        os.makedirs(directory)
    if os.path.exists(directory):
        # app.Documents.Item(1).ExportFile(idEPSType, myExportedEPSFile)
        myDocument.Export(idEPSType, myExportedEPSFile)

myDocument.Close()