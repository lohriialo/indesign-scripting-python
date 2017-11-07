"""
Use the DocumentPreferences object to change the
dimensions and orientation of the document
"""

import win32com.client

app = win32com.client.Dispatch('InDesign.Application.CC.2017')

myDocument = app.Documents.Add()

idLandscape = 2003395685
try:
    myDocument.DocumentPreferences.PageHeight = "800pt"
    myDocument.DocumentPreferences.PageWidth = "600pt"
    myDocument.DocumentPreferences.PageOrientation = idLandscape
    myDocument.DocumentPreferences.PagesPerDocument = 16
except Exception as e:
    print(e)

# Save the file (fill in a valid file path on your system).
myFile = r'C:\ServerTestFiles\TestDocument.indd'

myDocument = myDocument.Save(myFile)
myDocument.Close()
