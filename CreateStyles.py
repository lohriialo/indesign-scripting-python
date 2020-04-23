__author__ = "CrudeRags"
__version__ = "1.0"

"""
Create Paragraph and Character Styles
"""

import win32com.client

app = win32com.client.Dispatch('InDesign.Application.CC.2019')

try:
# Paragraph Style 1
    myParagraphStyle = app.ParagraphStyles.Add()
    myParagraphStyle.name = "Tamil Basic Paragraph"
    myParagraphStyle.appliedFont = "Tamil Bible"
    myParagraphStyle.basedOn = '[No Paragraph Style]'
    myParagraphStyle.fontStyle = 'Plain'
    myParagraphStyle.paragraphJustification = 1886020709
    myParagraphStyle.pointSize = 12.0
except Exception as e:
    print(e)
    pass

# Paragraph Style 2
newPStyle = app.ParagraphStyles.Add()
newPStyle.name = "Verse Indent"
newPStyle.appliedFont = 'Tamil Bible'
newPStyle.leftIndent = 2.8
newPStyle.skew = 15.0

# Paragraph Style 3
titleStyle = app.ParagraphStyles.Add()
titleStyle.name = "Chapter Title"
titleStyle.appliedFont = 'Tamil Bible'
titleStyle.pointSize = 24.0
titleStyle.strokeWeight = 0.20

howToUse = """
How to apply style?

Get the paragraph you want 
(see Navigate.py for getting the paragraph)
and use the property `appliedParagraphStyle` 
to apply the styles you created
"""

# Paragraph Style 4
chapterNoStyle = app.ParagraphStyles.Add()
chapterNoStyle.name = "Chapter Number"
chapterNoStyle.appliedFont = 'Times New Roman'
chapterNoStyle.pointSize = 24.0
chapterNoStyle.strokeWeight = 0.20


# Character Style 1
englishStyle = app.CharacterStyles.Add()
englishStyle.name = "English"
englishStyle.appliedFont = 'Times New Roman'
englishStyle.pointSize = 12.0

