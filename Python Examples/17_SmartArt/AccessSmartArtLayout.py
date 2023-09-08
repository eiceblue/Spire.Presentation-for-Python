from spire.presentation.common import *
from spire.presentation import *


def AppendAllText(fname: str, text: List[str]):
    fp = open(fname, "w")
    for s in text:
        fp.write(s + "\n")
    fp.close()


inputFile = "./Data/SmartArt.pptx"
outputFile = "AccessSmartArtLayout.txt"

# Create PPT document
presentation = Presentation()

# Load the PPT
presentation.LoadFromFile(inputFile)
strB = []


for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, ISmartArt):
        # Get the SmartArt and collect nodes
        sa = shape if isinstance(shape, ISmartArt) else None
        # Check SmartArt Layout
        layout = str(sa.LayoutType)
        strB.append("SmartArt layout type is " + layout)

AppendAllText(outputFile, strB)
presentation.Dispose()
