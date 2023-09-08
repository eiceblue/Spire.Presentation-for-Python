from spire.presentation.common import *
from spire.presentation import *

inputFile = "./Data/RemoveNoteFromSlides.pptx"
outputFile = "RemoveNotesAtSpecificSlide.pptx"

# Create a PPT document
presentation = Presentation()

# Load PPT file from disk
presentation.LoadFromFile(inputFile)
# Get the first slide
slide = presentation.Slides[0]

# Get note slide
note = slide.NotesSlide
# Clear note text
note.NotesTextFrame.Text = ""

# Save the PPT to PDF file format
presentation.SaveToFile(outputFile, FileFormat.Pptx2007)
presentation.Dispose()
