from spire.presentation.common import *
import math
from spire.presentation import *



outputFile ="CreateSlide.pptx"
#Create PPT document
presentation = Presentation()
#Add new slide
presentation.Slides.Append()
#Set the background image
for i in range(0, 2):
    ImageFile = "Data/bg.png"
    rect = RectangleF.FromLTRB (0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height)
    presentation.Slides[i].Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, ImageFile, rect)
    presentation.Slides[i].Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.get_FloralWhite()
#Add title
left =math.trunc(presentation.SlideSize.Size.Width / float(2)) - 200
rec_title = RectangleF.FromLTRB (left, 70, 400+left, 120)
shape_title = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec_title)
shape_title.ShapeStyle.LineColor.Color = Color.get_White()
shape_title.Fill.FillType = FillFormatType.none
para_title = TextParagraph()
para_title.Text = "E-iceblue"
para_title.Alignment = TextAlignmentType.Center
para_title.TextRanges[0].LatinFont = TextFont("Myriad Pro Light")
para_title.TextRanges[0].FontHeight = 36
para_title.TextRanges[0].Fill.FillType = FillFormatType.Solid
para_title.TextRanges[0].Fill.SolidColor.Color = Color.get_Black()
shape_title.TextFrame.Paragraphs.Append(para_title)
#Append new shape
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (50, 150, 650, 430))
shape.ShapeStyle.LineColor.Color = Color.get_White()
shape.Fill.FillType = FillFormatType.none
shape.Line.FillType = FillFormatType.none
#Add text to shape
shape.AppendTextFrame("Welcome to use Spire.Presentation for .NET.")
#Add new paragraph
pare = TextParagraph()
pare.Text = ""
shape.TextFrame.Paragraphs.Append(pare)
#Add new paragraph
pare = TextParagraph()
pare.Text = "Spire.Presentation for .NET is a professional PowerPoint compatible component that enables developers to create, read, write, modify, convert and Print PowerPoint documents from any .NET(C#, VB.NET, ASP.NET) platform. As an independent PowerPoint .NET component, Spire.Presentation for .NET doesn't need Microsoft PowerPoint installed on the machine."
shape.TextFrame.Paragraphs.Append(pare)
#Set the Font
for para in shape.TextFrame.Paragraphs:
    para.TextRanges[0].LatinFont = TextFont("Myriad Pro")
    para.TextRanges[0].FontHeight = 24
    para.TextRanges[0].Fill.FillType = FillFormatType.Solid
    para.TextRanges[0].Fill.SolidColor.Color = Color.get_Black()
    para.Alignment = TextAlignmentType.Left
#Append new shape - SixPointedStar
shape = presentation.Slides[1].Shapes.AppendShape(ShapeType.SixPointedStar, RectangleF.FromLTRB (100, 100, 200, 200))
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_Orange()
shape.ShapeStyle.LineColor.Color = Color.get_White()
#Append new shape
shape = presentation.Slides[1].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (50, 250, 650, 300))
shape.ShapeStyle.LineColor.Color = Color.get_White()
shape.Fill.FillType = FillFormatType.none
#Add text to shape
shape.AppendTextFrame("This is newly added Slide.")
#Set the Font
shape.TextFrame.Paragraphs[0].TextRanges[0].LatinFont = TextFont("Myriad Pro")
shape.TextFrame.Paragraphs[0].TextRanges[0].FontHeight = 24
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.FillType = FillFormatType.Solid
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.get_Black()
shape.TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Left
shape.TextFrame.Paragraphs[0].Indent = 35
#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()