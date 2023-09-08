from spire.presentation.common import *
from spire.presentation import *


def AppendAllText(fname: str, text: List[str]):
    fp = open(fname, "w")
    for s in text:
        fp.write(s + "\n")
    fp.close()


inputFile = "Data/SmartArt.pptx"
outputFile = "output/AccessChildNode.txt"

# Create PPT document
presentation = Presentation()

# Load the PPT
presentation.LoadFromFile(inputFile)

strB = []
strB.append("Access SmartArt child nodes.")
strB.append("Here is the SmartArt child node parameters details:")
outString = ""


for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, ISmartArt):
        # Get the SmartArt and collect nodes
        sa = shape if isinstance(shape, ISmartArt) else None
        nodes = sa.Nodes

        position = 0
        # Access the parent node at position 0
        node = nodes[position]
        childnode = None
        # Traverse through all child nodes inside SmartArt
        for i, node in enumerate(node.ChildNodes):
            # Access SmartArt child node at index i
            childnode = node
            # Print the SmartArt child node parameters
            outString = "Node text = "+childnode.TextFrame.Text+", Node level = " + \
                str(childnode.Level)+", Node Position = " + \
                str(childnode.Position)
            strB.append(outString)

# Save the file
AppendAllText(outputFile, strB)
presentation.Dispose()
