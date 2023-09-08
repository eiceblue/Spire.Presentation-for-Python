from spire.presentation.common import *
from spire.presentation import *


def AppendAllText(fname: str, text: List[str]):
    fp = open(fname, "w")
    for s in text:
        fp.write(s + "\n")
    fp.close()


inputFile = "Data/SmartArt.pptx"
outputFile = "output/AccessSpecificChildNode.txt"

# Create PPT document
presentation = Presentation()

# Load the PPT
presentation.LoadFromFile(inputFile)

strB = []
strB.append("Access SmartArt child node at specific position.")
strB.append("Here is the SmartArt child node parameters details:")
for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, ISmartArt):
        # Get the SmartArt
        sa = shape if isinstance(shape, ISmartArt) else None

        # Get SmartArt node collection
        nodes = sa.Nodes

        # Access SmartArt node at index 0
        node = nodes[0]

        # Access SmartArt child node at index 1
        childNode = node.ChildNodes[1]

        # Print the SmartArt child node parameters
        outString = "Node text = "+childNode.TextFrame.Text+", Node level = " + \
            str(childNode.Level)+", Node Position = "+str(childNode.Position)

        strB.append(outString)

    # Save the file
AppendAllText(outputFile, strB)
presentation.Dispose()
