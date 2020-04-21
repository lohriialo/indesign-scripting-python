# NavigateDocument.py

# Created: 4/17/2020 

__author__ = "CrudeRags"
__version__ = "1.1"

"""
Navigate a book, and its documents pagewise and storywise
"""

import win32com.client
import os

#Use your version of InDesign here
app = win32com.client.Dispatch('InDesign.Application.CC.2019')  
idnPath = os.path.abspath(r"path_to_book")
bookPath = os.path.join(idnPath,'MH 1-43.indb')

# ShowingWindow if false would not show what is opened in the app. If set to true, the book/document/library will be opened in the app
app.Open(From = bookPath,  ShowingWindow = False)
myBook = app.ActiveBook

# bookContents property gives access to bookContent object. The specific content has to opened separately
for doc in myBook.bookContents:
    doc_name = doc.name
    #Open document
    myDoc = app.Open(From = doc.fullName, ShowingWindow=False)
    #Get first story from document
    doc_story = myDoc.Stories[0]
    #Navigate a document page wise
    for myPage in myDoc.Pages:     
        #Get the text frames in the page
        for myFrame in myPage.TextFrames:
            #Get Contents directly: Type = str
            myContents = myFrame.Contents
            
            #Get paragraphs in the text frame
            for myPara in myFrame.Paragraphs:
                #Get Paragraph style
                print(myPara.appliedParagraphStyle)
                # if str(myPara.appliedParagraphStyle) == "Basic Paragraph":
                    # do something

            #Navigate a document storywise
    for story in myDoc.Stories:
        # print(story.Contents) 

        for para in story.Paragraphs:
            #do stuff
            para.Contents += "Hi there!"


            #Get all the paragraph styles in the document
    for style in myDoc.ParagraphStyles:
        print(style)