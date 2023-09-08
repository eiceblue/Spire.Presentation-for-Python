from spire.presentation.common import *
from spire.presentation import *

inputFile = "./Data/Template_Ppt_5.pptx"
outputFile = "RemoveHyperlink.pptx"

# Create a PowerPoint document.
presentation = Presentation()

# Load the file from disk.
presentation.LoadFromFile(inputFile)

# Get the shape and its text with hyperlink.
shape = presentation.Slides[0].Shapes[0] if isinstance(
    presentation.Slides[0].Shapes[0], IAutoShape) else None

# Set the ClickAction property into null to remove the hyperlink.
shape.TextFrame.TextRange.ClickAction = None

# Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()
