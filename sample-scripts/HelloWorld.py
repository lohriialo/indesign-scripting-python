# HelloWorld.jsx

import win32com.client
import os

app = win32com.client.Dispatch('InDesign.Application.CC.2017')

# Get reference to InDesign application
# app = indesign.Dispatch('InDesign.Application.CC.2017')
# Create a new document.
myDocument = app.Documents.Add()
# Get a reference to the first page.
myPage = myDocument.Pages.Item(1)
# Create a text frame.
myTextFrame = myPage.TextFrames.Add()
# Specify the size and shape of the text frame.
myTextFrame.GeometricBounds = ["6p0", "6p0", "18p0", "18p0"]
# Enter text in the text frame.
myTextFrame.Contents = "Hello World!"
# Save the document (fill in a valid file path).
myFile = r'C:\ServerTestFiles\HelloWorld.indd'
directory = os.path.dirname(myFile)
try:
    # If file path does not exist, create directory
    if not os.path.exists(directory):
        os.makedirs(directory)
    # If file path exist, save the file to the path
    if os.path.exists(directory):
        myDocument.Save(myFile)
except Exception as e:
    print('Export to PDF failed: ' + str(e))
# Close the document.
myDocument.Close()
