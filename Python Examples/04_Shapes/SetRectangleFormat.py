from spire.presentation import *
import math

outputFile = "SetRectangleFormat.pptx"

#Create a PPT document
presentation = Presentation()
#Add a shape
left =math.trunc(presentation.SlideSize.Size.Width / float(2)) - 100
rect = RectangleF.FromLTRB (left, 100, 200+left, 200)
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rect)
#Set the fill format of shape
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_CadetBlue()
#Set the fill format of line
shape.Line.FillType = FillFormatType.Solid
shape.Line.SolidFillColor.Color = Color.get_DimGray()
#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()