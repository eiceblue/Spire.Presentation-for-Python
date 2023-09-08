from spire.presentation.common import *
from spire.presentation import *


def AppendAllText(fname: str, text: List[str]):
    fp = open(fname, "w")
    for s in text:
        fp.write(s + "\n")
    fp.close()


inputFile = "Data/SmartArt.pptx"
outputFile = "output/AccessSmartArt.txt"

# Create PPT document
presentation = Presentation()

# Load the PPT
presentation.LoadFromFile(inputFile)

strB = []
strB.append("Access SmartArt nodes.")
strB.append("Here is the SmartArt node parameters details:")
outString = ""
node = None
for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, ISmartArt):
        # Get the SmartArt
        sa = shape if isinstance(shape, ISmartArt) else None

        nodes = sa.Nodes

        # Traverse through all nodes inside SmartArt
        for i, unusedItem in enumerate(nodes):
            # Access SmartArt node at index i
            node = nodes[i]
            # Print the SmartArt node parameters
            outString = "Node text = "+node.TextFrame.Text+", Node level = " + \
                str(node.Level)+", Node Position = " + \
                str(node.Position)
            strB.append(outString)

# Save the file
AppendAllText(outputFile, strB)
presentation.Dispose()
