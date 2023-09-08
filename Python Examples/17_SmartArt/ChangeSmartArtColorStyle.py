from spire.presentation.common import *
from spire.presentation import *


inputFile = "./Data/AddSmartArtNode.pptx"
outputFile = "ChangeSmartArtColorStyle.pptx"

# Create PPT document
presentation = Presentation()

# Load the PPT
presentation.LoadFromFile(inputFile)

for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, ISmartArt):
        # Get the SmartArt and collect nodes
        smartArt = shape if isinstance(shape, ISmartArt) else None
        # Check SmartArt color type
        if smartArt.ColorStyle == SmartArtColorType.ColoredFillAccent1:
            # Change SmartArt color type
            smartArt.ColorStyle = SmartArtColorType.ColorfulAccentColors

# Save the file
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()
