"""
CloseAllDocument.py
Closes all open documents without saving.
"""

import win32com.client

app = win32com.client.Dispatch('InDesign.Application.CC.2017')

idNo = 1852776480
for x in range(0, app.Documents.Count):
    app.Documents.Item(1).Close(idNo)
