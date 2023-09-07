from spire.presentation.common import *
from spire.presentation import *


outputFile ="AddLineToSlide.pptx"
#Create a PPT document
presentation = Presentation()
#Get the first slide
slide = presentation.Slides[0]
#Add a line in the slide
line = slide.Shapes.AppendShape(ShapeType.Line, RectangleF.FromLTRB (50, 100, 350, 100))
#Set color of the line
line.ShapeStyle.LineColor.Color = Color.get_Red()
#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()