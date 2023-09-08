from spire.presentation.common import *
import math
from spire.presentation import *

inputFile = "./Data/InsertVideo.pptx"
outputFile = "InsertVideo.pptx"

# Create a PPT document
presentation = Presentation()

# Load the document from disk
presentation.LoadFromFile(inputFile)

# Add title
rec_title = RectangleF.FromLTRB(50, 280, 160+50, 50+280)
shape_title = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.Rectangle, rec_title)
shape_title.ShapeStyle.LineColor.Color = Color.get_Transparent()

shape_title.Fill.FillType = FillFormatType.none
para_title = TextParagraph()
para_title.Text = "Video:"
para_title.Alignment = TextAlignmentType.Center
para_title.TextRanges[0].LatinFont = TextFont("Myriad Pro Light")
para_title.TextRanges[0].FontHeight = 32
para_title.TextRanges[0].IsBold = TriState.TTrue
para_title.TextRanges[0].Fill.FillType = FillFormatType.Solid
para_title.TextRanges[0].Fill.SolidColor.Color = Color.FromArgb(
    255, 68, 68, 68)
shape_title.TextFrame.Paragraphs.Append(para_title)

# Insert video into the document
left = math.trunc(presentation.SlideSize.Size.Width / float(2)) - 125
videoRect = RectangleF.FromLTRB(left, 240, 150+left, 150+240)
video = presentation.Slides[0].Shapes.AppendVideoMedia(
    "Data/Video.mp4", videoRect)
video.PictureFill.Picture.Url = "Data/Video.png"

# Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()
