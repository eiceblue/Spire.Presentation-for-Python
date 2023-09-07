from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/GetShapeGroupAltText.pptx"
outputFile ="GetShapeAltText.txt"
#Create a PPT document
presentation = Presentation()
#Load document from disk
presentation.LoadFromFile(inputFile)
builder = []
#Loop through slides and shapes
for slide in presentation.Slides:
    for shape in slide.Shapes:
        if isinstance(shape, GroupShape):
            #Find the shape group
            groupShape = shape if isinstance(shape, GroupShape) else None
            for gShape in groupShape.Shapes:
                #Append the alternative text in builder
                builder.append (gShape.AlternativeText)
#Write the content in txt file
fp = open(outputFile,"w")
for s in builder:
    fp.write(s + "\n")
presentation.Dispose()