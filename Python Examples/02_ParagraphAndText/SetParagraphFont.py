from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/Template_Az2.pptx"
outputFile ="SetParagraphFont.pptx"

#Create a PPT document
presentation = Presentation()

#Load PPT file from disk
presentation.LoadFromFile(inputFile)
#Get the first slide
slide = presentation.Slides[0]

#Access the first and second placeholder in the slide and typecasting it as AutoShape
tf1 = (slide.Shapes[0]).TextFrame
tf2 = (slide.Shapes[1]).TextFrame

# Access the first Paragraph
para1 = tf1.Paragraphs[0]
para2 = tf2.Paragraphs[0]

#Justify the paragraph
para2.Alignment = TextAlignmentType.Justify

#Access the first text range
textRange1 = para1.FirstTextRange
textRange2 = para2.FirstTextRange

#Define new fonts
fd1 = TextFont("Elephant")
fd2 = TextFont("Castellar")

# Assign new fonts to text range
textRange1.LatinFont = fd1
textRange2.LatinFont = fd2

# Set font to Bold
textRange1.Format.IsBold = TriState.TTrue
textRange2.Format.IsBold = TriState.TFalse

# Set font to Italic
textRange1.Format.IsItalic = TriState.TFalse
textRange2.Format.IsItalic = TriState.TTrue

# Set font color
textRange1.Fill.FillType = FillFormatType.Solid
textRange1.Fill.SolidColor.Color = Color.get_Purple()
textRange2.Fill.FillType = FillFormatType.Solid
textRange2.Fill.SolidColor.Color = Color.get_Peru()

presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()


