from spire.presentation.common import *
from spire.presentation import *

# Create a Presentation object
presentation = Presentation()

# Load a Presentation document
presentation.LoadFromFile("AddSmartArtNode.pptx")  

sb = []

# Get custom document properties
for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, ISmartArt):
        sa = shape
        sb.append(f"left: {sa.Left}")
        sb.append(f"top: {sa.Top}")

with open("out.txt", 'a') as f:
    for line in sb:
        f.write(line + '\n')

presentation.Dispose()