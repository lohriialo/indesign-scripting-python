"""
Use the DocumentPreferences object to change the
dimensions and orientation of the document
"""

import win32com.client

app = win32com.client.Dispatch('InDesign.Application.CC.2017')

myFile = r'C:\ServerTestFiles\TestDocument.indd'
myDocument = app.Open(myFile)

myDocumentPrefs = myDocument.DocumentPreferences

try:
    myDocumentPreset = app.DocumentPresets.Item("myDocumentPreset")
    presetName = myDocumentPreset.Name
    print('preset already exist: ' + presetName)
except Exception as e:
    print(e)
    myDocumentPreset = app.DocumentPresets.Add()
    myDocumentPreset.Name = "myDocumentPreset"
    myDocumentPreset.Left = myDocument.MarginPreferences.Left
    myDocumentPreset.Right = myDocument.MarginPreferences.Right
    myDocumentPreset.Top = myDocument.MarginPreferences.Top
    myDocumentPreset.Bottom = myDocument.MarginPreferences.Bottom
    myDocumentPreset.ColumnCount = myDocument.MarginPreferences.ColumnCount
    myDocumentPreset.ColumnGutter = myDocument.MarginPreferences.ColumnGutter
    myDocumentPreset.DocumentBleedBottomOffset = myDocumentPrefs.DocumentBleedBottomOffset
    myDocumentPreset.DocumentBleedTopOffset = myDocumentPrefs.DocumentBleedTopOffset
    myDocumentPreset.DocumentBleedInsideOrLeftOffset = myDocumentPrefs.DocumentBleedInsideOrLeftOffset
    myDocumentPreset.DocumentBleedOutsideOrRightOffset = myDocumentPrefs.DocumentBleedOutsideOrRightOffset
    myDocumentPreset.FacingPages = myDocument.DocumentPreferences.FacingPages
    myDocumentPreset.PageHeight = myDocument.DocumentPreferences.PageHeight
    myDocumentPreset.PageWidth = myDocument.DocumentPreferences.PageWidth
    myDocumentPreset.PageOrientation = myDocument.DocumentPreferences.PageOrientation
    myDocumentPreset.PagesPerDocument = myDocument.DocumentPreferences.PagesPerDocument
    myDocumentPreset.SlugBottomOffset = myDocument.DocumentPreferences.SlugBottomOffset
    myDocumentPreset.SlugTopOffset = myDocument.DocumentPreferences.SlugTopOffset
    myDocumentPreset.SlugInsideOrLeftOffset = myDocument.DocumentPreferences.SlugInsideOrLeftOffset
    myDocumentPreset.SlugRightOrOutsideOffset = myDocument.DocumentPreferences.SlugRightOrOutsideOffset
