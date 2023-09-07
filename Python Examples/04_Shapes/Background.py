from spire.presentation.common import *
import math

from spire.presentation import *

outputFile ="Background.pptx"
#Create a PPT document
presentation = Presentation()
#Set background Image
ImageFile = "./Data/backgroundImg.png"
rect = RectangleF.FromLTRB(0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height)
presentation.Slides[0].Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, ImageFile, rect)
#Add title
left  = math.trunc(presentation.SlideSize.Size.Width / float(2)) - 200
rec_title = RectangleF.FromLTRB (left, 70, 380+left, 120)
shape_title = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec_title)
shape_title.Line.FillType = FillFormatType.none
shape_title.Fill.FillType =FillFormatType.none
para_title = TextParagraph()
para_title.Text = "Background Sample"
para_title.Alignment = TextAlignmentType.Center
para_title.TextRanges[0].LatinFont = TextFont("Lucida Sans Unicode")
para_title.TextRanges[0].FontHeight = 36
para_title.TextRanges[0].Fill.FillType = FillFormatType.Solid
para_title.TextRanges[0].Fill.SolidColor.Color = Color.get_DarkSlateBlue()
shape_title.TextFrame.Paragraphs.Append(para_title)
#Add new shape to PPT document
left = math.trunc(presentation.SlideSize.Size.Width / float(2)) - 300
rec = RectangleF.FromLTRB (left, 155, 600+left, 355)
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec)
shape.Line.FillType = FillFormatType.none
shape.Fill.FillType = FillFormatType.none
para = TextParagraph()
para.Text = "Spire.Presentation for .NET support PPT, PPS, PPTX and PPSX presentation formats. It provides functions such as managing text, image, shapes, tables, animations, audio and video on slides. It also support exporting presentation slides to EMF, JPG, TIFF, PDF format etc."
para.TextRanges[0].Fill.FillType = FillFormatType.Solid
para.TextRanges[0].Fill.SolidColor.Color = Color.get_CadetBlue()
para.TextRanges[0].FontHeight = 26
shape.TextFrame.Paragraphs.Append(para)
#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()