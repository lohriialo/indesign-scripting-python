"""
Add a series of guides using the createGuides method
"""

import win32com.client

app = win32com.client.Dispatch('InDesign.Application.CC.2017')

myDocument = app.Documents.Add()

"""
Parameters (all optional): row count, column count, row gutter,
column gutter,guide color, fit margins, remove existing, layer.
Note that the createGuides method does not take an RGB array
for the guide color parameter.
"""
idGray = 1766290041
myDocument.Spreads.Item(1).CreateGuides(4, 4, "1p", "1p", idGray, True, True, myDocument.Layers.Item(1))

# Save the file (fill in a valid file path on your system).
myFile = r'C:\ServerTestFiles\TestDocument.indd'

myDocument = myDocument.Save(myFile)
myDocument.Close()
