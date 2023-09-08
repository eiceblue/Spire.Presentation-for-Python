from spire.presentation.common import *
from spire.presentation import *


inputFile = "Data/AddSmartArtNode.pptx"
outputFile = "output/ChangeSmartArtShapeStyle.pptx"

# Create PPT document
presentation = Presentation()

# Load the PPT
presentation.LoadFromFile(inputFile)

for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, ISmartArt):
        # Get the SmartArt and collect nodes
        smartArt = shape if isinstance(shape, ISmartArt) else None
        # Check SmartArt style
        if smartArt.Style == SmartArtStyleType.SimpleFill:
            # Change SmartArt Style
            smartArt.Style = SmartArtStyleType.Cartoon

# Save the file
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()
