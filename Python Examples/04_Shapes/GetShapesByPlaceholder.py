from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/GetShapesByPlaceholder.pptx"
outputFile ="GetShapesByPlaceholder.txt"      
ppt = Presentation()
ppt.LoadFromFile(inputFile)
placeholder = ppt.Slides[1].Shapes[0].Placeholder
#Get Shapes by Placeholder
shapes = ppt.Slides[1].GetPlaceholderShapes(placeholder)
text = ""
#Iterate over all the shapes
i = 0
while i < len(shapes):
    #If shape is IAutoShape
    if isinstance(shapes[i], IAutoShape):
        autoShape = shapes[i] if isinstance(shapes[i], IAutoShape) else None
        if autoShape.TextFrame is not None:
            text += autoShape.TextFrame.Text + "\r\n"
    i += 1
#Save to file.
fp = open(outputFile,"w")
for s in text:
    fp.write(s)
fp.close()
ppt.Dispose()