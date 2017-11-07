"""
Exports a range of pages to a PDF file.
Enter the names of the pages you want to export in the following line.
Note that the page name is not necessarily the index of the page in the
document (e.g., the first page of a document whose page numbering starts
with page 21 will be "21", not 1).
"""

import win32com.client
import os

app = win32com.client.Dispatch('InDesign.Application.CC.2017')

idPDFType = 1952403524
idPageOrigin = 1380143215
idAutoPageNumber = 1396797550
idCenterAlign = 1667591796
idAscentOffset = 1296135023

myDocument = app.Documents.Add()
myDocument.ViewPreferences.RulerOrigin = idPageOrigin
myDocument.DocumentPreferences.PagesPerDocument = 12
myMasterSpread = myDocument.MasterSpreads.Item(1)

def myGetBounds(myDocument, myPage):
    myPageWidth = myDocument.DocumentPreferences.PageWidth
    myPageHeight = myDocument.DocumentPreferences.PageHeight
    myMarginPreferences = myPage.MarginPreferences
    myLeft = myMarginPreferences.Left
    myTop = myMarginPreferences.Top
    myRight = myPageWidth - myMarginPreferences.Right
    myBottom = myPageHeight - myMarginPreferences.Bottom
    return [myTop, myLeft, myBottom, myRight]

for x in range(0, myMasterSpread.Pages.Count):
    myTextFrame = myMasterSpread.Pages.Item(x + 1).TextFrames.Add()
    myTextFrame.Move(None, [1, 1])
    myTextFrame.Contents = idAutoPageNumber
    myTextFrame.Paragraphs.Item(1).PointSize = 72
    myTextFrame.Paragraphs.Item(1).Justification = idCenterAlign
    myTextFrame.TextFramePreferences.FirstBaselineOffset = idAscentOffset
    myTextFrame.TextFramePreferences.VerticalJustification = idCenterAlign
    myTextFrame.GeometricBounds = myGetBounds(myDocument, myMasterSpread.Pages.Item(x + 1))

app.PDFExportPreferences.PageRange = "1-3, 6, 9"
myExportedPDFFile = r'C:\ServerTestFiles\TestDocument.pdf'
directory = os.path.dirname(myExportedPDFFile)

if not os.path.exists(directory):
    os.makedirs(directory)
if os.path.exists(directory):
    myDocument.Export(idPDFType, myExportedPDFFile)

