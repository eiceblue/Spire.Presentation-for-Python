from spire.presentation import *

inputFile = "./Data/OverlappingShapes.pptx"
outputFile = "ReorderOverlappingShapes_out.pptx"

#Create an instance of presentation document
ppt = Presentation()
#Load file
ppt.LoadFromFile(inputFile)
#Get the first shape of the first slide
shape = ppt.Slides[0].Shapes[0]
#Change the shape's zorder
ppt.Slides[0].Shapes.ZOrder(1, shape)
#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()