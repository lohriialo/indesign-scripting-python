"""
OpenDocument.py
Opens an existing document. You'll have to fill in your own file path
"""

import win32com.client
import os

app = win32com.client.Dispatch('InDesign.Application.CC.2017')

myFile = r'C:\ServerTestFiles\TestDocument.indd'
directory = os.path.dirname(myFile)

if not os.path.exists(directory):
    os.makedirs(directory)
    myDocument = app.Documents.Add()
    myDocument = myDocument.Save(myFile)
    myDocument.Close()
if os.path.exists(directory):
    myDocument = app.Open(myFile)
    myPage = myDocument.Pages.Item(1)
    myRectangle = myPage.Rectangles.Add()
    myRectangle.GeometricBounds = ["6p", "6p", "18p", "18p"]
    myRectangle.StrokeWeight = 12
    #leave the document open...
    print(myDocument.FullName)
    # myDocument.Close()
