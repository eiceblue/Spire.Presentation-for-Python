from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/FindShapeByAltText.pptx"
outputFile ="FindShapeByAltText.txt"
#Create a PPT document
presentation = Presentation()
#Load document from disk
presentation.LoadFromFile(inputFile)
#Get the first slide
slide = presentation.Slides[0]
#Find shape in the slide
for shape in slide.Shapes:
    #Find the shape whose alternative text is altText
    if shape.AlternativeText=="Shape1":
        fp = open(outputFile,"w")
        fp.write(str(shape.Name) + "\n")
        fp.close()
          
presentation.Dispose()



