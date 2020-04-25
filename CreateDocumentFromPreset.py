__author__ = "CrudeRags"
__version__ = "1.0"

"""
Create a document using a preset
"""

import win32com.client

app = win32com.client.Dispatch('InDesign.Application.CC.2019')
my_file = r'C:\ServerTestFiles\TestDocument.indd'

# See all the available local presets
for preset in app.DocumentPresets:
    print(preset.name) 

# Choose one from the above presets. If you want your own preset, see CreateDocumentPreset.py
# Once you create a document preset it will persist. No need to create it every time
doc_preset = app.DocumentPresets.Item("7x9Book")
myDoc = app.Documents.Add(DocumentPreset = doc_preset)
myDoc.Save()
myDoc.Close()

