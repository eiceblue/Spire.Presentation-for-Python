from spire.presentation.common import *
from spire.presentation import *

inputFile = "./Data/ExtractOLEObject.pptx"
outputFile_px = "ExtractOLEObject.pptx"
outputFile_p = "ExtractOLEObject.ppt"
outputFile_xls = "ExtractOLEObject.xls"
outputFile_xlsx = "ExtractOLEObject.xlsx"
outputFile_doc = "ExtractOLEObject.doc"
outputFile_docx = "ExtractOLEObject.docx"

# Create a PPT document
presentation = Presentation()

# Load document from disk
presentation.LoadFromFile(inputFile)

# Loop through the slides and shapes
for slide in presentation.Slides:
    for shape in slide.Shapes:
        if isinstance(shape, IOleObject):
            # Find OLE object
            oleObject = shape if isinstance(shape, IOleObject) else None

            # Get its data and write to file
            stream = oleObject.Data
            if oleObject.ProgId == "Excel.Sheet.8":
                stream.Save(outputFile_xls)
            elif oleObject.ProgId == "Excel.Sheet.12":
                stream.Save(outputFile_xlsx)
            elif oleObject.ProgId == "Word.Document.8":
                stream.Save(outputFile_doc)
            elif oleObject.ProgId == "Word.Document.12":
                stream.Save(outputFile_docx)
            elif oleObject.ProgId == "PowerPoint.Show.8":
                stream.Save(outputFile_p)
            elif oleObject.ProgId == "PowerPoint.Show.12":
                stream.Save(outputFile_px)
            stream.Dispose()
presentation.Dispose()
