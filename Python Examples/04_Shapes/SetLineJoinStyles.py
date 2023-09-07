from spire.presentation import *

outputFile = "SetLineJoinStyles.pptx"

#Create a PPT document
presentation = Presentation()
#Get the first slide
slide = presentation.Slides[0]
#Add three shapes
shape1 = slide.Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (50, 150, 200, 200))
shape2 = slide.Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (250, 150, 400, 200))
shape3 = slide.Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (450, 150, 600, 200))
#Fill shapes
shape1.Fill.FillType = FillFormatType.Solid
shape1.Fill.SolidColor.Color = Color.get_CadetBlue()
shape2.Fill.FillType = FillFormatType.Solid
shape2.Fill.SolidColor.Color = Color.get_CadetBlue()
shape3.Fill.FillType = FillFormatType.Solid
shape3.Fill.SolidColor.Color = Color.get_CadetBlue()
#Fill lines of shapes
shape1.Line.FillType = FillFormatType.Solid
shape1.Line.SolidFillColor.Color = Color.get_DarkGray()
shape2.Line.FillType = FillFormatType.Solid
shape2.Line.SolidFillColor.Color = Color.get_DarkGray()
shape3.Line.FillType = FillFormatType.Solid
shape3.Line.SolidFillColor.Color = Color.get_DarkGray()
#Set the line width
shape1.Line.Width = 10
shape2.Line.Width = 10
shape3.Line.Width = 10
#Set the join styles of lines
shape1.Line.JoinStyle = LineJoinType.Bevel
shape2.Line.JoinStyle = LineJoinType.Miter
shape3.Line.JoinStyle = LineJoinType.Round
#Add text in shapes
shape1.TextFrame.Text = "Bevel Join Style"
shape2.TextFrame.Text = "Miter Join Style"
shape3.TextFrame.Text = "Round Join Style"
#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()