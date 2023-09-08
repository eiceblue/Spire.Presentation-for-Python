from spire.presentation.common import *
from spire.presentation import *


inputFile = "Data/AddSmartArtNode2.pptx"
outputFile = "output/AddNodeByPosition.pptx"

# Create PPT document
presentation = Presentation()

# Load the PPT
presentation.LoadFromFile(inputFile)

for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, ISmartArt):
        # Get the SmartArt and collect nodes
        smartArt = shape if isinstance(shape, ISmartArt) else None

        position = 0
        # Add a new node at specific position
        node = smartArt.Nodes.AddNodeByPosition(position)
        # Add text and set the text style
        node.TextFrame.Text = "New Node"
        node.TextFrame.TextRange.Fill.FillType = FillFormatType.Solid
        node.TextFrame.TextRange.Fill.SolidColor.KnownColor = KnownColors.Red

        # Get a node
        node = smartArt.Nodes[1]
        position = 1
        # Add a new child node at specific position
        childNode = node.ChildNodes.AddNodeByPosition(position)
        # Add text and set the text style
        node.TextFrame.Text = "New child node"
        node.TextFrame.TextRange.Fill.FillType = FillFormatType.Solid
        node.TextFrame.TextRange.Fill.SolidColor.KnownColor = KnownColors.Blue


# Save the file
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()
