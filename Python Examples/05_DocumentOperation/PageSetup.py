from spire.presentation import *
import math

inputFile = "./Data/bg.png"
outputFile = "PageSetup_out.pptx"

#Create PPT document
presentation = Presentation()
#Set the size of slides
size = SizeF(600.0,600.0)
presentation.SlideSize.Size = size
presentation.SlideSize.Orientation = SlideOrienation.Portrait
presentation.SlideSize.Type = SlideSizeType.Custom
#Set background image
rect = RectangleF.FromLTRB (0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height)
presentation.Slides[0].Shapes.AppendEmbedImageByPath(ShapeType.Rectangle, inputFile, rect)
presentation.Slides[0].Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.get_FloralWhite()
#Append new shape
left = math.trunc(presentation.SlideSize.Size.Width / float(2)) - 200
rec = RectangleF.FromLTRB (left, 150, 400+left, 350)
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec)
shape.ShapeStyle.LineColor.Color = Color.get_White()
shape.Fill.FillType = FillFormatType.none
#Add text to shape
shape.AppendTextFrame("The sample demonstrates how to set slide size.")
shape.TextFrame.Paragraphs[0].TextRanges[0].LatinFont = TextFont("Myriad Pro")
shape.TextFrame.Paragraphs[0].TextRanges[0].FontHeight = 24
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.FillType = FillFormatType.Solid
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.FromArgb(255,36, 64, 97)
#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()