"""
Creates a new document preset.
If the document preset "7x9Book" does not already exist, create it.
"""

import win32com.client

app = win32com.client.Dispatch('InDesign.Application.CC.2017')

idPortrait = 1751738216
try:
    myDocumentPreset = app.DocumentPresets.Item("7x9Book")
    presetName = myDocumentPreset.Name
    print('preset already exist')
except Exception as e:
    print(e)
    myDocumentPreset = app.DocumentPresets.Add()
    myDocumentPreset.Name = "7x9Book"

    myDocumentPreset.PageHeight = "9i"
    myDocumentPreset.PageWidth = "7i"
    myDocumentPreset.Left = "4p"
    myDocumentPreset.Right = "6p"
    myDocumentPreset.Top = "4p"
    myDocumentPreset.Bottom = "9p"
    myDocumentPreset.ColumnCount = 1
    myDocumentPreset.DocumentBleedBottomOffset = "3p"
    myDocumentPreset.DocumentBleedTopOffset = "3p"
    myDocumentPreset.DocumentBleedInsideOrLeftOffset = "3p"
    myDocumentPreset.DocumentBleedOutsideOrRightOffset = "3p"
    myDocumentPreset.FacingPages = True
    myDocumentPreset.PageOrientation = idPortrait
    myDocumentPreset.PagesPerDocument = 1
    myDocumentPreset.SlugBottomOffset = "18p"
    myDocumentPreset.SlugTopOffset = "3p"
    myDocumentPreset.SlugInsideOrLeftOffset = "3p"
    myDocumentPreset.SlugRightOrOutsideOffset = "3p"
