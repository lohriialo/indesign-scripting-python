"""
Opens an InDesign Template, iterate text frames by script label and replace text from user input
Exports to PDF
"""
from flask import Flask, render_template, request, redirect, url_for, send_from_directory
import win32com.client
import os
import pythoncom

app = Flask(__name__)

businessCardTemplate = r'C:\ServerTestFiles\BusinessCardTemplate.indt'
myBusinessCard = 'myBusinessCard.pdf'
businessCardFullPath = r'C:\ServerTestFiles' + '\\' + myBusinessCard
directory = os.path.dirname(businessCardFullPath)

@app.route('/', methods=['GET'])
def my_form():
    return render_template("index.html")

@app.route('/downloadPDF/<path:filename>')
def downloadPDF(filename):
    return send_from_directory(directory, filename, as_attachment=False)

@app.route('/processData', methods=['GET', 'POST'])
def processData():
    """
    Why CoInitialize? See this for details
    https://stackoverflow.com/questions/26745617/win32com-client-dispatch-cherrypy-coinitialize-has-not-been-called
    """
    pythoncom.CoInitialize()

    indesign = win32com.client.Dispatch('InDesignServer.Application.CC.2017')
    myDocument = indesign.Open(businessCardTemplate)

    textFrames = myDocument.TextFrames

    firstName = request.form['firstName']
    lastName = request.form['lastName']
    jobTitle = request.form['title']
    jobEmail = request.form['email']
    for x in range(textFrames.Count):
        if textFrames[x].Label == 'first':
            textFrames[x].Contents = firstName
        if textFrames[x].Label == 'last':
            textFrames[x].Contents = lastName
        if textFrames[x].Label == 'title':
            textFrames[x].Contents = jobTitle
        if textFrames[x].Label == 'email':
            textFrames[x].Contents = jobEmail

    idPDFType = 1952403524
    idNo = 1852776480
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
        if os.path.exists(directory):
            myDocument.Export(idPDFType, businessCardFullPath)
            myDocument.Close(idNo)
            return render_template("download.html", pdf=myBusinessCard)
    except Exception as e:
        return 'Export to PDF failed: ' + str(e)

if __name__ == '__main__':
    app.run(debug=True)
