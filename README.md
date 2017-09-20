Scripting in InDesign is used to automate a wide variety of repetative task or as complex as an entire new feature. If you've never done scripting in InDesign, you should first read the scripting documentations, detail at https://console.adobe.io/downloads/id

InDesign can be scripted through COM(Component Object Model). Its DOM(Document Object Model) is the same when accessing it through either its own JavaScript engine or Python or any other scripting language. InDesign exposes it's scripting DOM as a Type Library file at the time of InDesign application startup. This is because all the scripting methods available in the DOM is provided by the individual InDesign Scriptable Plug-Ins. Those Plug-Ins could be the stock InDesign Plug-Ins that ships with InDesign or any other third party Plug-Ins which as installed in the Plug-Ins directory. The type library file is written out during application launch at C:\Users\username\AppData\Local\Adobe\InDesign\<Version>\en_US\Caches\Scripting Support\<Version>\Resources for Visual Basic.tlb

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
