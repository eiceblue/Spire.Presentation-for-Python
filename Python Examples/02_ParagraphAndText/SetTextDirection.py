from spire.presentation.common import *
from spire.presentation import *


outputFile ="SetTextDirection.pptx"

#Create an instance of presentation document
ppt = Presentation()

#Append a shape with text to the first slide
textboxShape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (250, 70, 350, 470))
textboxShape.ShapeStyle.LineColor.Color = Color.get_Transparent()
textboxShape.Fill.FillType = FillFormatType.Solid
textboxShape.Fill.SolidColor.Color = Color.get_LightBlue()
textboxShape.TextFrame.Text = "You Are Welcome Here"
#Set the text direction to vertical
textboxShape.TextFrame.VerticalTextType = VerticalTextType.Vertical

#Append another shape with text to the slide
textboxShape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (350, 70, 450, 470))
textboxShape.ShapeStyle.LineColor.Color = Color.get_Transparent()
textboxShape.Fill.FillType = FillFormatType.Solid
textboxShape.Fill.SolidColor.Color = Color.get_LightGray()
#Append some asian characters
textboxShape.TextFrame.Text = "欢迎光临"
#Set the VerticalTextType as EastAsianVertical to aviod rotating text 90 degrees
textboxShape.TextFrame.VerticalTextType = VerticalTextType.EastAsianVertical

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()