# InDesign Scripting In Python
![](https://i.imgur.com/mnmAIDH.png "Logo Title Text 1")

Scripting in InDesign is used to automate a wide variety of repetative task or as complex as an entire new feature. It is one of the most powerful features in InDesign. It can save lots of time: almost anything you can do with the user interface, you can do with scripts, and there are even a few things you can do in scripting that you cannot do from within the user interface. Scripting is also a good way to do accurate placing and to repeat the same action dozens, hundreds, or thousands of times. If you've never done scripting in InDesign, you should first read the scripting [documentations](https://console.adobe.io/downloads/id)

# But why Python?
InDesign scripting officially supports JavaScript, AppleScript & VBScript and the Scripting SDK ships with tutorials and guides for each of those supported Scripting language. However, Scripting in Python is also fairly easy if not easier if you're already comfortable with Python. You may have already heard that Python is gaining in popularity, but did you know it’s now the most popular introductory programming language in U.S. universities? Python is also cross platform just like JavaScript is and lately becoming one of the fastest growing programming language according to StackOverflow [as of 2017](https://stackoverflow.blog/2017/09/06/incredible-growth-python) / [as of 2019](https://insights.stackoverflow.com/survey/2019#key-results)

Python is easy to use, powerful, and versatile, making it a great choice for beginners and experts alike. Python’s readability makes it a great first programming language - it allows you to think like a programmer and not waste time understanding the mysterious syntax that other programming languages can require.

# InDesign COM & DOM
InDesign can be scripted through COM(Component Object Model). Its DOM(Document Object Model) is the same when accessing it through either its own JavaScript engine or Python or any other scripting language. InDesign exposes it's scripting DOM as a Type Library file at the time of application startup. This is because all the scripting methods available in the DOM is provided by the individual scriptable Plug-Ins. Those Plug-Ins could be the stock Plug-Ins that ships with InDesign or any other third party Plug-Ins which as installed in the Plug-Ins directory. The type library file is written out during application launch at `C:\Users\username\AppData\Local\Adobe\InDesign\<Version>\en_US\Caches\Scripting Support\<Version>\Resources for Visual Basic.tlb`

Python allows you to access COM and it's DOM with the help of a Python extension called  "Python Win32 Extensions", for more details check https://sourceforge.net/projects/pywin32/

However, in order to install this extension, you have to manually download and link to your existing Python installation which can be cumbersome. Instead, there is now a version of pywin32 on PyPI that can be installed with pip. It is called pypiwin32, and it installs the package using the binary wheel format.

* `pip install pypiwin32`

Now to call InDesign COM

```python
import win32com.client
app = win32com.client.Dispatch('InDesign.Application.CC.2017')

myDocument = app.Documents.Add()
myPage = myDocument.Pages.Item(1)
myTextFrame = myPage.TextFrames.Add()
myTextFrame.GeometricBounds = ["6p0", "6p0", "18p0", "18p0"]
myTextFrame.Contents = "Hello World!"
```
# How to inspect scripting object properties?
There's not a straight forward way, you need to read the documentation to understand what properties/attributes are available for a scripting object, or possibly a COM browser. For example, I've extracted the Python scripting object reference for InDesign CC 2018 at [api_reference](https://github.com/lohriialo/indesign-scripting-python/tree/master/api_reference)

# More Example
[Here's](https://github.com/lohriialo/indesign-scripting-python/tree/master/flask-demo) another example of how easy it is to setup a web application with Flask in Python. Where, InDesign Server running on a server takes the user inputed data and generate a business cards on the fly. See [`flask-demo`](https://github.com/lohriialo/indesign-scripting-python/tree/master/flask-demo) for the example

# Scripting on Mac?
Yes, scripting on Mac is also possible, see [photoshop_mac_scripting](https://github.com/lohriialo/photoshop-scripting-python/tree/master/mac_scripting) for more details as a reference to get started

# Note
Python is not directly supported as a first party scripting language, so you cannot drop a python file in the scripts folder and have InDesign execute it. At the moment, the only way to script is via COM as demonstrated in the examples

There's a feature request though to add support for Python in InDesign, you might want to go give your vote [here](https://indesign.uservoice.com/forums/601021-adobe-indesign-feature-requests/suggestions/32193772-add-python-to-the-list-of-supported-scripting-lang)

# InDesign Scripting Resources
* [InDesign Scripting SDK](https://developer.adobe.com)
* [InDesign Scripting Documentation](https://developer.adobe.com)
* [InDesign Scripting Developer Forum](https://community.adobe.com/t5/indesign/ct-p/ct-indesign?page=1&sort=latest_replies&filter=all&lang=all&tabid=discussions&topics=label-scripting%2Clabel-sdk)
* [InDesign Scripting API Reference](https://www.indesignjs.de/extendscriptAPI/indesign-latest)

# Also see 
* [Photoshop Scripting in Python](https://github.com/lohriialo/photoshop-scripting-python)
* [Illustrator Scripting in Python](https://github.com/lohriialo/illustrator-scripting-python)

# Contribution
If you've written a useful InDesign Python script and wants to share with the world, please create a new issue with the file as an attachment to the issue.

When you submit a script, please try to include the following information at the start of your script
```python
# script_file_name.py

# Created: 1st January 2019
__author__ = 'Your Name or Original Author Name'
__version__ = '1.0'

"""
A short description of what the script does
"""

"""
Instructions on how to use the script, if any
"""

```
* Go to  [indesign-scripting-python/issues/new](https://github.com/lohriialo/indesign-scripting-python/issues/new)
* Add title  as `Useful Script`
* Drag & drop your .py script file into the description area
* Click `Submit new issue`
