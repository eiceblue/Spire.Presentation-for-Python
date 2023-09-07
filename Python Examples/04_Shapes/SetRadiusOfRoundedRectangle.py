from spire.presentation import *

outputFile = "SetRadiusOfRoundedRectangle.pptx"

#Create a PPT document
presentation = Presentation()
#Insert a rounded rectangle and set its radious
presentation.Slides[0].Shapes.InsertRoundRectangle(0, 160, 180, 100, 200, 10)
#Append a rounded rectangle and set its radius
shape = presentation.Slides[0].Shapes.AppendRoundRectangle(380, 180, 100, 200, 100)
#Set the color and fill style of shape
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_SeaGreen()
shape.ShapeStyle.LineColor.Color = Color.get_White()
#Rotate the shape to 90 degree
shape.Rotation = 90
#Save the document to Pptx file
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()