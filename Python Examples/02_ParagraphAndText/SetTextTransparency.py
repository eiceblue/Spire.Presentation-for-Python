from spire.presentation.common import *
from spire.presentation import *


outputFile ="SetTextTransparency.pptx"

#Create an instance of presentation document
ppt = Presentation()

#Set background image
ImageFile = "./Data/bg.png"
rect = RectangleF.FromLTRB (0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height)
ppt.Slides[0].Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, ImageFile, rect)
ppt.Slides[0].Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.get_FloralWhite()

#Add a shape
textboxShape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (100, 100, 400, 220))
textboxShape.ShapeStyle.LineColor.Color = Color.get_Transparent()
textboxShape.Fill.FillType = FillFormatType.none

#Remove default blank paragraphs
textboxShape.TextFrame.Paragraphs.Clear()

#Add three paragraphs, apply color with different alpha values to text
alpha = 55
for i in range(0, 3):
    textboxShape.TextFrame.Paragraphs.Append(TextParagraph())
    textboxShape.TextFrame.Paragraphs[i].TextRanges.Append(TextRange("Text Transparency"))
    textboxShape.TextFrame.Paragraphs[i].TextRanges[0].Fill.FillType = FillFormatType.Solid
    textboxShape.TextFrame.Paragraphs[i].TextRanges[0].Fill.SolidColor.Color = Color.FromArgb(alpha, Color.get_Purple().R,Color.get_Purple().G,Color.get_Purple().B)
    alpha += 100

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()

