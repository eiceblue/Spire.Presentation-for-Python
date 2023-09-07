from spire.presentation.common import *
from spire.presentation import *

inputImageFile = "./Data/bg.png"
outputFile ="AddParagraph.pptx"

#Create an instance of presentation document
ppt = Presentation()


#Set background image
rect = RectangleF.FromLTRB(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height)
ppt.Slides[0].Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, inputImageFile, rect)
ppt.Slides[0].Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.get_FloralWhite()

#Append a new shape
shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (50, 70, 670, 220))
shape.Fill.FillType = FillFormatType.none
shape.ShapeStyle.LineColor.Color = Color.get_White()

#Set the alignment of paragraph
shape.TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Left
#Set the indent of paragraph
shape.TextFrame.Paragraphs[0].Indent = 50
#Set the linespacing of paragraph
shape.TextFrame.Paragraphs[0].LineSpacing = 150
#Set the text of paragraph
shape.TextFrame.Text = "This powerful component suite contains the most up-to-date versions of all python components offered by E-iceblue."

#Set the Font
shape.TextFrame.Paragraphs[0].TextRanges[0].LatinFont = TextFont("Arial Rounded MT Bold")
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.FillType = FillFormatType.Solid
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.get_Black()

#Save and launch the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()