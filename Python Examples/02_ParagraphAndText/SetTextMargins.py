from spire.presentation.common import *
from spire.presentation import *


outputFile ="SetTextMargins.pptx"

#Create an instance of presentation document
ppt = Presentation()

#Set background image
ImageFile = "./Data/bg.png"
rect = RectangleF.FromLTRB (0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height)
ppt.Slides[0].Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, ImageFile, rect)
ppt.Slides[0].Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.get_FloralWhite()

#Append a new shape
shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (50, 100, 500, 250))

#Set margins for text inside shapes
shape.Fill.FillType = FillFormatType.none
shape.ShapeStyle.LineColor.Color = Color.get_LightBlue()
shape.TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Justify
shape.TextFrame.Text = "Spire.Presentation for Python is a professional presentation processing API that is highly compatible with PowerPoint. It is a completely independent class library that developers can use to create, edit, convert, and save PowerPoint presentations efficiently without installing Microsoft PowerPoint."
shape.TextFrame.Paragraphs[0].TextRanges[0].LatinFont = TextFont("Arial Rounded MT Bold")
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.FillType = FillFormatType.Solid
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.get_Black()

#Set the margins for the text frame
shape.TextFrame.MarginTop = 10
shape.TextFrame.MarginBottom = 35
shape.TextFrame.MarginLeft = 15
shape.TextFrame.MarginRight = 30

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()