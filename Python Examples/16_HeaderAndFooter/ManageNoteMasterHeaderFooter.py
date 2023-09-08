from spire.presentation.common import *
from spire.presentation import *

inputFile = "./Data/PPTHasHeader.pptx"
outputFile = "ManageNoteMasterHeaderFooter.pptx"

# Create a PPT document
presentation = Presentation()
# Load presentation
presentation.LoadFromFile(inputFile)

# Set the note Masters header and footer
noteMasterSlide = presentation.NotesMaster
if noteMasterSlide is not None:
    for shape in noteMasterSlide.Shapes:
        if shape.Placeholder is not None:
            if shape.Placeholder.Type is PlaceholderType.Header:
                (shape if isinstance(shape, IAutoShape)
                 else None).TextFrame.Text = "change the header by Spire"
            if shape.Placeholder.Type is PlaceholderType.Footer:
                (shape if isinstance(shape, IAutoShape)
                 else None).TextFrame.Text = "change the footer by Spire"

# Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()
