from spire.presentation.common import *
import math
from spire.presentation import *


inputFile ="./Data/Template_Az.pptx"
outputFile ="MultipleParagraphs.pptx"

#Create a PPT document
presentation = Presentation()

#Load PPT file from disk
presentation.LoadFromFile(inputFile)
#Access the first slide
slide = presentation.Slides[0]

# Add an AutoShape of rectangle type
left = math.trunc(presentation.SlideSize.Size.Width / float(2)) - 250
rec = RectangleF.FromLTRB (left, 150, 500+left, 300)
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec)

# Access TextFrame of the AutoShape
tf = shape.TextFrame

# Create Paragraphs and TextRanges with different text formats
para0 = tf.Paragraphs[0]
textRange1 = TextRange()
textRange2 = TextRange()
para0.TextRanges.Append(textRange1)
para0.TextRanges.Append(textRange2)

para1 = TextParagraph()
tf.Paragraphs.Append(para1)
textRange11 = TextRange()
textRange12 = TextRange()
textRange13 = TextRange()
para1.TextRanges.Append(textRange11)
para1.TextRanges.Append(textRange12)
para1.TextRanges.Append(textRange13)

para2 = TextParagraph()
tf.Paragraphs.Append(para2)
textRange21 = TextRange()
textRange22 = TextRange()
textRange23 = TextRange()
para2.TextRanges.Append(textRange21)
para2.TextRanges.Append(textRange22)
para2.TextRanges.Append(textRange23)

for i in range(0, 3):
    for j in range(0, 3):
        tf.Paragraphs[i].TextRanges[j].Text = "TextRange " + str(j)
        if j == 0:
            tf.Paragraphs[i].TextRanges[j].Fill.FillType = FillFormatType.Solid
            tf.Paragraphs[i].TextRanges[j].Fill.SolidColor.Color = Color.get_LightBlue()
            tf.Paragraphs[i].TextRanges[j].Format.IsBold = TriState.TTrue
            tf.Paragraphs[i].TextRanges[j].FontHeight = 15
        elif j == 1:
            tf.Paragraphs[i].TextRanges[j].Fill.FillType = FillFormatType.Solid
            tf.Paragraphs[i].TextRanges[j].Fill.SolidColor.Color = Color.get_Blue()
            tf.Paragraphs[i].TextRanges[j].Format.IsItalic = TriState.TTrue
            tf.Paragraphs[i].TextRanges[j].FontHeight = 18


presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()
