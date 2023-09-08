from spire.presentation.common import *
from spire.presentation import *


inputFile = "Data/AddSmartArtNode.pptx"
outputFile = "output/AddSmartArtNode.pptx"

# Create a PPT document
presentation = Presentation()

# Load the document from disk
presentation.LoadFromFile(inputFile)

# Get the SmartArt
sa = presentation.Slides[0].Shapes[0] if isinstance(
    presentation.Slides[0].Shapes[0], ISmartArt) else None

# Add a node
node = sa.Nodes.AddNode()
# Add text and set the text style
node.TextFrame.Text = "AddText"
node.TextFrame.TextRange.Fill.FillType = FillFormatType.Solid
node.TextFrame.TextRange.Fill.SolidColor.KnownColor = KnownColors.HotPink

presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()
