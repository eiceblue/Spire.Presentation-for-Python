from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/Template_Az1.pptx"
outputFile ="RotateText.pptx"

#Create a PPT document
presentation = Presentation()

#Load PPT file from disk
presentation.LoadFromFile(inputFile)
#Get the first slide
slide = presentation.Slides[0]
#Get a shape 
shape = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IAutoShape) else None

shape.TextFrame.VerticalTextType = VerticalTextType.Vertical270

presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()

