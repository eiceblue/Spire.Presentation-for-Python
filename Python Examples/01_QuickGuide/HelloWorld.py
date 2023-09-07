from spire.presentation.common import *
import math
from spire.presentation import *


outputFile ="HelloWorld.pptx"

#Create a PPT document
presentation = Presentation()

#Add a new shape to the PPT document
left = math.trunc(presentation.SlideSize.Size.Width / float(2))-250
rec = RectangleF.FromLTRB(left, 80, left+500, 150+80)
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec)
shape.ShapeStyle.LineColor.Color = Color.get_White()
shape.Fill.FillType = FillFormatType.none

#Add text to the shape
shape.AppendTextFrame("Hello World!")

#Set the font and fill style of the text
textRange = shape.TextFrame.TextRange
textRange.Fill.FillType = FillFormatType.Solid
textRange.Fill.SolidColor.Color = Color.get_CadetBlue()
textRange.FontHeight = 66
textRange.LatinFont = TextFont("Lucida Sans Unicode")

#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()