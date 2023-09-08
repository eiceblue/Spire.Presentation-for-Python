from spire.presentation.common import *
from spire.presentation import *

inputFile = "./Data/Template_Ppt_5.pptx"
outputFile = "RemoveSpeakerNotes.pptx"

# Create a PowerPoint document.
presentation = Presentation()

# Load the file from disk.
presentation.LoadFromFile(inputFile)

# Get the first slide from the sample document.
slide = presentation.Slides[0]

# Remove the first speak note.
slide.NotesSlide.NotesTextFrame.Paragraphs.RemoveAt(1)

# Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)

presentation.Dispose()
