"""
Exports the current document as PDF.
Assumes you have a document open.
document.exportFile parameters are:
Format: use either the ExportFormat.pdfType constant or the string "Adobe PDF"
To: a file path as a string
Using: PDF export preset (or a string that is the name of a PDF export preset)
The default PDF export preset names are surrounded by square brackets (e.g., "[Screen]").
"""

import win32com.client
import os

app = win32com.client.Dispatch('InDesign.Application.CC.2017')

myInddFile = r'C:\ServerTestFiles\TestDocument.indd'
myDocument = app.Open(myInddFile)

myPDFFile = r'C:\ServerTestFiles\TestDocument.pdf'
directory = os.path.dirname(myPDFFile)

# def setattrs(_self, **kwargs):
#     for k, v in kwargs.items():
#         setattr(_self, k, v)
#
# pdfExportPreferences = app.PDFExportPreferences
# pdfExportPreferences.setattr(pdfExportPreferences,
#         ExportGuidesAndGrids=False,
#         ExportLayers=False,
#         ExportNonPrintingObjects=False,
#         ExportReaderSpreads=False,
#         GenerateThumbnails=False
#         )

try:
    if not os.path.exists(directory):
        os.makedirs(directory)
    if os.path.exists(directory):
        try:
            pdfExportPreferences = app.PDFExportPreferences
            # Basic PDF output options.
            idAllPages = 1886547553
            idAcrobat6 = 1097020976
            pdfExportPreferences.PageRange = idAllPages
            pdfExportPreferences.AcrobatCompatibility = idAcrobat6
            pdfExportPreferences.ExportGuidesAndGrids = False
            pdfExportPreferences.ExportLayers = False
            #pdfExportPreferences.ExportNonPrintingObjects = False
            pdfExportPreferences.ExportReaderSpreads = False
            pdfExportPreferences.GenerateThumbnails = False
            try:
                pdfExportPreferences.IgnoreSpreadOverrides = False
            except Exception as e:
                print("IgnoreSpreadOverrides: " + str(e))
            pdfExportPreferences.IncludeBookmarks = True
            pdfExportPreferences.IncludeHyperlinks = True
            pdfExportPreferences.IncludeICCProfiles = True
            pdfExportPreferences.IncludeSlugWithPDF = False
            pdfExportPreferences.IncludeStructure = False
            #pdfExportPreferences.InteractiveElements = False
            """
            Setting subsetFontsBelow to zero disallows font subsetting
            set subsetFontsBelow to some other value to use font subsetting.
            """
            pdfExportPreferences.SubsetFontsBelow = 0
            # Bitmap compression/sampling/quality options
            idZip = 2053730371
            idEightBit = 1701722210
            idNone = 1852796517
            pdfExportPreferences.ColorBitmapCompression = idZip
            pdfExportPreferences.ColorBitmapQuality = idEightBit
            pdfExportPreferences.ColorBitmapSampling = idNone

            # thresholdToCompressColor is not needed in this example
            # colorBitmapSamplingDPI is not needed when colorBitmapSampling is set to none
            pdfExportPreferences.GrayscaleBitmapCompression = idZip
            pdfExportPreferences.GrayscaleBitmapQuality = idEightBit
            pdfExportPreferences.GrayscaleBitmapSampling = idNone

            # thresholdToCompressGray is not needed in this example
            # grayscaleBitmapSamplingDPI is not needed when grayscaleBitmapSampling is set to none
            pdfExportPreferences.MonochromeBitmapCompression = idZip
            pdfExportPreferences.MonochromeBitmapSampling = idNone
            # thresholdToCompressMonochrome is not needed in this example
            # monochromeBitmapSamplingDPI is not needed when monochromeBitmapSampling is set to none

            # Other compression options
            idCompressNone = 1131368047
            pdfExportPreferences.CompressionType = idCompressNone
            pdfExportPreferences.CompressTextAndLineArt = True
            pdfExportPreferences.CropImagesToFrames = True
            pdfExportPreferences.OptimizePDF = True

            # Printers marks and prepress options.
            # Get the bleed amounts from the document's bleed.
            bleedBottom = app.Documents.Item(1).DocumentPreferences.DocumentBleedBottomOffset
            bleedTop = app.Documents.Item(1).DocumentPreferences.DocumentBleedTopOffset
            bleedInside = app.Documents.Item(1).DocumentPreferences.DocumentBleedInsideOrLeftOffset
            bleedOutside = app.Documents.Item(1).DocumentPreferences.DocumentBleedOutsideOrRightOffset
            # If any bleed area is greater than zero, then export the bleed marks.
            if bleedBottom is 0 and bleedTop is 0 and bleedInside is 0 and bleedOutside is 0:
                pdfExportPreferences.BleedMarks = True
            else:
                pdfExportPreferences.BleedMarks = False

            # Default mark type
            idDefault = 1147563124
            pdfExportPreferences.PDFMarkType = idDefault
            idP125pt = 825374064
            printerMarkWeight = idP125pt
            pdfExportPreferences.RegistrationMarks = True
            try:
                pdfExportPreferences.SimulateOverprint = False
            except Exception as e:
                print("SimulateOverprint: " + str(e))
            pdfExportPreferences.UseDocumentBleedWithPDF = True
            # Set viewPDF to true to open the PDF in Acrobat or Adobe Reader
            pdfExportPreferences.ViewPDF = True

        except Exception as e:
            print(e)
        # Now do the export
        idPDFType = 1952403524
        myDocument.Export(idPDFType, myPDFFile)
except Exception as e:
    print('Export to PDF failed: ' + str(e))
myDocument.Close()
