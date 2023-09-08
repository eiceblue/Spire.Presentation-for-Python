from spire.presentation.common import *
from spire.presentation import *

inputFile = "./Data/CreateSmartArtShape.pptx"
outputFile = "CreateSmartArtShape.pptx"

# Create a PPT document
presentation = Presentation()

# Load the document from disk
presentation.LoadFromFile(inputFile)

sa = presentation.Slides[0].Shapes.AppendSmartArt(
    200, 60, 300, 300, SmartArtLayoutType.Gear)

# Set type and color of smartart
sa.Style = SmartArtStyleType.SubtleEffect
sa.ColorStyle = SmartArtColorType.GradientLoopAccent3

# Remove all shapes
to_remove = []


for a in sa.Nodes:
    to_remove.append(a)
for subnode in to_remove:
    sa.Nodes.RemoveNode(subnode)

    # Add two custom shapes with text
node = sa.Nodes.AddNode()
sa.Nodes[0].TextFrame.Text = "aa"
node = sa.Nodes.AddNode()
node.TextFrame.Text = "bb"
node.TextFrame.TextRange.Fill.FillType = FillFormatType.Solid
node.TextFrame.TextRange.Fill.SolidColor.KnownColor = KnownColors.Black

# Save and launch the file
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()
