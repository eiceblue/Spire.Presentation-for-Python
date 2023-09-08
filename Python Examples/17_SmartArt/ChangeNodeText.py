from spire.presentation.common import *
from spire.presentation import *


inputFile = "Data/AddSmartArtNode.pptx"
outputFile = "output/ChangeNodeText.pptx"

# Create PPT document
presentation = Presentation()

# Load the PPT
presentation.LoadFromFile(inputFile)

for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, ISmartArt):
        # Get the SmartArt and collect nodes
        smartArt = shape if isinstance(shape, ISmartArt) else None
        # Obtain the reference of a node by using its Index
        # select second root node
        node = smartArt.Nodes[1]
        # Set the text of the TextFrame
        node.TextFrame.Text = "Second root node"
# Save the file
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()
