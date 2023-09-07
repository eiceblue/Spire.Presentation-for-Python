from spire.presentation import *

inputFile = "./Data/RotateShape.pptx"
outputFile = "RotateShape_out.pptx"

#Load a PPT document
ppt = Presentation()
ppt.LoadFromFile(inputFile)
#Get the shapes 
shape = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IAutoShape) else None
#Set the rotation
shape.Rotation = 60
( ppt.Slides[0].Shapes[1] if isinstance(ppt.Slides[0].Shapes[1], IAutoShape) else None).Rotation = 120
( ppt.Slides[0].Shapes[2] if isinstance(ppt.Slides[0].Shapes[2], IAutoShape) else None).Rotation = 180
( ppt.Slides[0].Shapes[3] if isinstance(ppt.Slides[0].Shapes[3], IAutoShape) else None).Rotation = 240
#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2010)
ppt.Dispose()