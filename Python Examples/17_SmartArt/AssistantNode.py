from spire.presentation.common import *
from spire.presentation import *


inputFile = "Data/AddSmartArtNode.pptx"
outputFile = "output/AssistantNode.pptx"

# Create PPT document
presentation = Presentation()

# Load the PPT
presentation.LoadFromFile(inputFile)
node = None
for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, ISmartArt):
        # Get the SmartArt and collect nodes
        smartArt = shape if isinstance(shape, ISmartArt) else None

        nodes = smartArt.Nodes

        # Traverse through all nodes inside SmartArt
        for i, unusedItem in enumerate(nodes):
            # Access SmartArt node at index i
            node = nodes[i]
            # Check if node is assitant node
            if not node.IsAssistant:
                # Set node as assitant node
                node.IsAssistant = True
# Save the file
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()
