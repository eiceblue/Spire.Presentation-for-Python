from spire.presentation.common import *
from spire.presentation import *




outputFile ="AddLineWithArrow.pptx"
#Create an instance of presentation document
ppt = Presentation()
#Set background image
ImageFile = "Data/bg.png"
rect = RectangleF.FromLTRB (0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height)
ppt.Slides[0].Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, ImageFile, rect)
ppt.Slides[0].Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.get_FloralWhite()
#Add a line to the slides and set its color to red
shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Line, RectangleF.FromLTRB (150, 100, 250, 200))
shape.ShapeStyle.LineColor.Color = Color.get_Red()
#Set the line end type as StealthArrow
shape.Line.LineEndType = LineEndType.StealthArrow
#Add a line to the slides and use default color
shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Line, RectangleF.FromLTRB (300, 150, 400, 250))
shape.Rotation = -45
#Set the line end type as TriangleArrowHead
shape.Line.LineEndType = LineEndType.TriangleArrowHead
#Add a line to the slides and set its color to Green
shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Line, RectangleF.FromLTRB (450, 100, 550, 200))
shape.ShapeStyle.LineColor.Color = Color.get_Green()
shape.Rotation = 90
#Set the line begin type as TriangleArrowHead
shape.Line.LineBeginType = LineEndType.StealthArrow
#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()