from spire.presentation.common import *
import math

from spire.presentation import *


outputFile ="FillShapeWithSolidColor.pptx"
#Create a PPT document
presentation = Presentation()
#Get the first slide
slide = presentation.Slides[0]
#Add a rectangle
left = math.trunc(presentation.SlideSize.Size.Width / float(2)) - 50
rect = RectangleF.FromLTRB (left, 100, 100+left, 200)
shape = slide.Shapes.AppendShape(ShapeType.Rectangle, rect)
#Fill shape with solid color
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_Yellow()
#Set the fill format of line
shape.Line.FillType = FillFormatType.Solid
shape.Line.SolidFillColor.Color = Color.get_Gray()
#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()