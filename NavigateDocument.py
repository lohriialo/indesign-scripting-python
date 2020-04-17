# NavigateDocument.py

# Created: 4/17/2020 

__author__ = "CrudeRags"
__version__ = "1.0"

"""
Navigate a document pagewise and storywise
"""

import win32com.client
import os

#Use your version of InDesign here
app = win32com.client.Dispatch('InDesign.Application.CC.2019')  
myFile = r'C:\ServerTestFiles\TestDocument.indd'

#Open document
myDocument = app.Open(myFile)

#Navigate a document page wise
for pageIndex in range(myDocument.Pages.Count):     
    #Page Reference - It gives a handle to manipulate page
    myPage = myDocument.Pages.Item(pageIndex)   # myDocument.Pages[pageIndex] also works - list style indexing can be substitued wherever Item is used
    #Get the text frames in the page
    for frameIndex in range(1,myPage.TextFrames.Count):
        #Get the text frame reference
        myFrame = myPage.TextFrames.Item(frameIndex)
        #Get Contents directly: Type = str
        myContents = myFrame.Contents
        
        #Get paragraphs in the text frame
        for para_index in range(myFrame.Paragraphs.Count):
            myPara = myFrame.Paragraphs[para_index]
            #Get Paragraph style
            print(myPara.appliedParagraphStyle)
            # if str(myPara.appliedParagraphStyle) == "Basic Paragraph":
                # do something

#Navigate a document storywise
for storyIndex in range(myDocument.Stories.Count):
    #Story handle
    myStory = myDocument.Stories[storyIndex]
    

#Get all the paragraph styles in the document
for style in myDocument.ParagraphStyles:
    print(style)