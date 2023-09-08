from spire.presentation.common import *
import math
from spire.presentation import *

outputFile = "Hyperlinks.pptx"


#Create a PPT document
presentation = Presentation()

#Set background Image
ImageFile = "./Data/bg.png"
rect = RectangleF.FromLTRB (0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height)
presentation.Slides[0].Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, ImageFile, rect)

#Add new shape to PPT document
left = math.trunc(presentation.SlideSize.Size.Width / float(2)) - 255
rec = RectangleF.FromLTRB (left, 120, 500+left, 400)
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec)
shape.Fill.FillType = FillFormatType.none
shape.Line.Width = 0

#Add some paragraphs with hyperlinks
para1 = TextParagraph()
tr = TextRange("E-iceblue")
tr.Fill.FillType = FillFormatType.Solid
tr.Fill.SolidColor.Color = Color.get_Blue()
para1.TextRanges.Append(tr)
para1.Alignment = TextAlignmentType.Center
shape.TextFrame.Paragraphs.Append(para1)
shape.TextFrame.Paragraphs.Append(TextParagraph())

#Add some paragraphs with hyperlinks
para2 = TextParagraph()
tr1 = TextRange("Click to know more about Spire.Presentation.")
tr1.ClickAction.Address = "http://www.e-iceblue.com/Introduce/presentation-for-net-introduce.html"
para2.TextRanges.Append(tr1)
shape.TextFrame.Paragraphs.Append(para2)
shape.TextFrame.Paragraphs.Append(TextParagraph())

para3 = TextParagraph()
tr2 = TextRange("Click to visit E-iceblue Home page.")
tr2.ClickAction.Address = "https://www.e-iceblue.com/"
para3.TextRanges.Append(tr2)
shape.TextFrame.Paragraphs.Append(para3)
shape.TextFrame.Paragraphs.Append(TextParagraph())

para4 = TextParagraph()
tr3 = TextRange("Click to go to the forum to raise questions.")
tr3.ClickAction.Address = "https://www.e-iceblue.com/forum/components-f5.html"
para4.TextRanges.Append(tr3)
shape.TextFrame.Paragraphs.Append(para4)
shape.TextFrame.Paragraphs.Append(TextParagraph())

para5 = TextParagraph()
tr4 = TextRange("Click to contact our sales team via email.")
tr4.ClickAction.Address = "mailto:sales@e-iceblue.com"
para5.TextRanges.Append(tr4)
shape.TextFrame.Paragraphs.Append(para5)
shape.TextFrame.Paragraphs.Append(TextParagraph())

para6 = TextParagraph()
tr5 = TextRange("Click to contact our support team via email.")
tr5.ClickAction.Address = "mailto:support@e-iceblue.com"
para6.TextRanges.Append(tr5)
shape.TextFrame.Paragraphs.Append(para6)

for para in shape.TextFrame.Paragraphs:
    if len(para.Text) != 0:
        para.TextRanges[0].LatinFont = TextFont("Lucida Sans Unicode")
        para.TextRanges[0].FontHeight = 20


#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()

     


    
