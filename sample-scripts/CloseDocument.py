"""
CloseAllDocument.py
Closes the current document.
"""

import win32com.client

app = win32com.client.Dispatch('InDesign.Application.CC.2017')

if app.Documents.Count is not 0:
    app.Documents.Item(1).Close()
else:
    print('No documents are open')
