from spire.presentation.common import *
from spire.presentation import *

inputFile = "./Data/Template_Ppt_5.pptx"
outputFile = "ModifyHyperlink.pptx"

# Create a PowerPoint document.
presentation = Presentation()

# Load the file from disk.
presentation.LoadFromFile(inputFile)

# Find the hyperlinks you want to edit.
shape = presentation.Slides[0].Shapes[0]

# Edit the link text and the target URL.
shape.TextFrame.TextRange.ClickAction.Address = "http://www.e-iceblue.com"
shape.TextFrame.TextRange.Text = "E-iceblue"

# Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()
