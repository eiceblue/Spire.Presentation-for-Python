from spire.presentation.common import *
from spire.presentation import *

inputFile = "Data/SmartArtLinklineOutline.pptx"
outputFile = "output/SetSmartArtLinklineOutline.pptx"

# Create a PPT document
ppt = Presentation()
# Load document from disk
ppt.LoadFromFile(inputFile)
smartArt = ppt.Slides[0].Shapes[0] if isinstance(
    ppt.Slides[0].Shapes[0], ISmartArt) else None
count = smartArt.Nodes.Count
node = None
# Loop through all smartArts
for i in range(0, count):
    node = smartArt.Nodes[i]
    # Set the line type
    node.LinkLine.FillType = FillFormatType.Solid
    # Set the line color
    node.LinkLine.SolidFillColor.Color = Color.get_Red()
    # Set the line width
    node.LinkLine.Width = 2
    # Set the line DashStyle
    node.LinkLine.DashStyle = LineDashStyleType.SystemDash

# Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()
