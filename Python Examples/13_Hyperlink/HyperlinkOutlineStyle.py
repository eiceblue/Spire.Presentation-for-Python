from spire.presentation.common import *
import math
from spire.presentation import *

outputFile = "HyperlinkOutlineStyle.pptx"


#Create a PPT document
presentation = Presentation()

#Add new shape to PPT document
left =math.trunc(presentation.SlideSize.Size.Width / float(2)) - 255
rec = RectangleF.FromLTRB (left, 120, 400+left, 220)
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec)
shape.Fill.FillType = FillFormatType.none
shape.Line.FillType = FillFormatType.none

#Add a paragraph with hyperlink
para1 = TextParagraph()
tr1 = TextRange("Click to know more about Spire.Presentation")
tr1.ClickAction.Address = "https://www.e-iceblue.com/Introduce/presentation-for-python.html"
para1.TextRanges.Append(tr1)

#Set the format of textrange
tr1.Format.FontHeight = 20
tr1.IsItalic = TriState.TTrue

#Set the outline format of textrange
tr1.TextLineFormat.FillFormat.FillType = FillFormatType.Solid
tr1.TextLineFormat.FillFormat.SolidFillColor.Color = Color.get_LightSeaGreen()
tr1.TextLineFormat.JoinStyle = LineJoinType.Round
tr1.TextLineFormat.Width = 2

#Add the paragraph to shape
shape.TextFrame.Paragraphs.Append(para1)
shape.TextFrame.Paragraphs.Append(TextParagraph())

#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()

     


    
