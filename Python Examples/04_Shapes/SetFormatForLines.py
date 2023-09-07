from spire.presentation import *

inputFile = "./Data/bg.png"
outputFile = "SetFormatForLines.pptx"

#Create an instance of presentation document
ppt = Presentation()
#Set background image
rect = RectangleF.FromLTRB (0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height)
ppt.Slides[0].Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, inputFile, rect)
ppt.Slides[0].Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.get_FloralWhite()
#Add a rectangle shape to the slide
shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (100, 150, 300, 250))
#Set the fill color of the rectangle shape
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_White()
#Apply some formatting on the line of the rectangle
shape.Line.Style = TextLineStyle.ThickThin
shape.Line.Width = 5
shape.Line.DashStyle = LineDashStyleType.Dash
#Set the color of the line of the rectangle
shape.ShapeStyle.LineColor.Color = Color.get_SkyBlue()
#Add a ellipse shape to the slide
shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Ellipse, RectangleF.FromLTRB (400, 150, 600, 250))
#Set the fill color of the ellipse shape
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_White()
#Apply some formatting on the line of the ellipse
shape.Line.Style = TextLineStyle.ThickBetweenThin
shape.Line.Width = 5
shape.Line.DashStyle = LineDashStyleType.DashDot
#Set the color of the line of the ellipse
shape.ShapeStyle.LineColor.Color = Color.get_OrangeRed()
#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()