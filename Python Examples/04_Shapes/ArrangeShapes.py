from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/ArrangeShape.pptx"
outputFile ="ArrangeShapes.pptx"

#Create an instance of presentation document
ppt = Presentation()
#Load file
ppt.LoadFromFile(inputFile)
#Get the specified shape
shape = ppt.Slides[0].Shapes[0]
#Bring the shape forward through SetShapeArrange method
shape.SetShapeArrange(ShapeArrange.BringForward)
#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()