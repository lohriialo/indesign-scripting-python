"""
Exports the current document as PDF.
Assumes you have a document open.
document.exportFile parameters are:
Format: use either the ExportFormat.pdfType constant or the string "Adobe PDF"
To: a file path as a string
Using: PDF export preset (or a string that is the name of a PDF export preset)
The default PDF export preset names are surrounded by square brackets (e.g., "[Screen]").
"""

import win32com.client
import os

app = win32com.client.Dispatch('InDesign.Application.CC.2017')

myInddFile = r'C:\ServerTestFiles\TestDocument.indd'
myDocument = app.Open(myInddFile)

myPDFFile = r'C:\ServerTestFiles\TestDocument.pdf'
directory = os.path.dirname(myPDFFile)

idPDFType = 1952403524
# 1=[High Quality Print], 2=[PDF/X-1a:2001] etc..
myPDFPreset = app.PDFExportPresets.Item(1)
try:
    if not os.path.exists(directory):
        os.makedirs(directory)
    if os.path.exists(directory):
        myDocument.Export(idPDFType, myPDFFile, False, myPDFPreset)
except Exception as e:
    print('Export to PDF failed: ' + str(e))
myDocument.Close()

