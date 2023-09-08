from spire.presentation.common import *
from spire.presentation import *


inputFile = "Data/RemoveNode.pptx"
outputFile = "output/RemoveNode.pptx"

# Create PPT document
presentation = Presentation()

# Load the document from disk
presentation.LoadFromFile(inputFile)

# Get the SmartArt and collect nodes
sa = presentation.Slides[0].Shapes[0] if isinstance(
    presentation.Slides[0].Shapes[0], ISmartArt) else None
nodes = sa.Nodes

# Remove the node to specific position
nodes.RemoveNodeByPosition(2)

presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()
