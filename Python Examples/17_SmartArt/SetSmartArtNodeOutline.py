from spire.presentation.common import *
from spire.presentation import *


inputFile = "./Data/SmartArtLinklineOutline.pptx"
outputFile = "SetSmartArtNodeOutline.pptx"

ppt = Presentation()
ppt.LoadFromFile(inputFile)
smartArt = ppt.Slides[0].Shapes[0] if isinstance(
    ppt.Slides[0].Shapes[0], ISmartArt) else None
count = smartArt.Nodes.Count
node = None
# Loop through all nodes
for i in range(0, count):
    node = smartArt.Nodes[i]
    # Set the fill format type
    node.Line.FillType = FillFormatType.Solid
    # Set the line style
    node.Line.Style = TextLineStyle.ThinThin
    # Set the line color
    node.Line.SolidFillColor.Color = Color.get_Red()
    # Set the line width
    node.Line.Width = 2

# Save to file.
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()
