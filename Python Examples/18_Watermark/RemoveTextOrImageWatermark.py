from spire.presentation.common import *
from spire.presentation import *


inputFile = "./Data/RemoveTextAndImageWatermarks.pptx"
outputFile = "RemoveTextAndImageWatermarks.pptx"

# Create a PowerPoint document.
presentation = Presentation()

# Load the file from disk.
presentation.LoadFromFile(inputFile)

# Remove text watermark by removing the shape which contains the text string "E-iceblue".
for i, unusedItem in enumerate(presentation.Slides):
    for j, unusedItem in enumerate(presentation.Slides[i].Shapes):
        if isinstance(presentation.Slides[i].Shapes[j], IAutoShape):
            shape = presentation.Slides[i].Shapes[j]
            if shape.TextFrame.Text.find("E-iceblue") != -1:
                presentation.Slides[i].Shapes.Remove(shape)

# Remove image watermark.
for i, unusedItem in enumerate(presentation.Slides):
    presentation.Slides[i].SlideBackground.Fill.FillType = FillFormatType.none

# Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()
