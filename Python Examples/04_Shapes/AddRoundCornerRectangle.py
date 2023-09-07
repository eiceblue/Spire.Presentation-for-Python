from spire.presentation.common import *
from spire.presentation import *


outputFile ="AddRoundCornerRectagle.pptx"
#Create an instance of presentation document
ppt = Presentation()
#Set background image
ImageFile = "Data/bg.png"
rect = RectangleF.FromLTRB (0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height)
ppt.Slides[0].Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, ImageFile, rect)
ppt.Slides[0].Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.get_FloralWhite()
#Append a round corner rectangle and set its radius
shape = ppt.Slides[0].Shapes.AppendRoundRectangle(300, 90, 100, 200, 80)
#Set the color and fill style of shape
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_LightBlue()
shape.ShapeStyle.LineColor.Color = Color.get_SkyBlue()
#Rotate the shape to 90 degree
shape.Rotation = 90
#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()