from spire.presentation.common import *
import math
from spire.presentation import *


outputFile ="SetTextFontProperties.pptx"

#Create a PPT document
presentation = Presentation()

#Add a new shape to the PPT document
left = math.trunc(presentation.SlideSize.Size.Width / float(2)) - 250
rec = RectangleF.FromLTRB (left, 80, 500+left, 230)
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec)

shape.ShapeStyle.LineColor.Color = Color.get_White()
shape.Fill.FillType = FillFormatType.none

#Add text to the shape
shape.AppendTextFrame("Welcome to use Spire.Presentation")

textRange = shape.TextFrame.TextRange
#Set the font
textRange.LatinFont = TextFont("Times New Roman")
#Set bold property of the font
textRange.IsBold = TriState.TTrue

#Set italic property of the font
textRange.IsItalic = TriState.TTrue

#Set underline property of the font
textRange.TextUnderlineType = TextUnderlineType.Single

#Set the height of the font
textRange.FontHeight = 50

#Set the color of the font
textRange.Fill.FillType = FillFormatType.Solid
textRange.Fill.SolidColor.Color = Color.get_CadetBlue()

presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()
