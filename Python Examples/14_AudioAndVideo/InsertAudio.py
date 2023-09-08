from spire.presentation.common import *
from spire.presentation import *

inputFile = "./Data/InsertAudio.pptx"
outputFile = "InsertAudio.pptx"

# Create a PPT document
presentation = Presentation()

# Load the document from disk
presentation.LoadFromFile(inputFile)

# Add title
rec_title = RectangleF.FromLTRB(50, 240, 160+50, 50+240)
shape_title = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.Rectangle, rec_title)
shape_title.ShapeStyle.LineColor.Color = Color.get_Transparent()

shape_title.Fill.FillType = FillFormatType.none
para_title = TextParagraph()
para_title.Text = "Audio:"
para_title.Alignment = TextAlignmentType.Center
para_title.TextRanges[0].LatinFont = TextFont("Myriad Pro Light")
para_title.TextRanges[0].FontHeight = 32
para_title.TextRanges[0].IsBold = TriState.TTrue
para_title.TextRanges[0].Fill.FillType = FillFormatType.Solid
para_title.TextRanges[0].Fill.SolidColor.Color = Color.FromArgb(
    255, 68, 68, 68)
shape_title.TextFrame.Paragraphs.Append(para_title)

# Insert audio into the document
audioRect = RectangleF.FromLTRB(220, 240, 80+220, 80+240)
presentation.Slides[0].Shapes.AppendAudioMedia(
    "./Data/Music.wav", audioRect)

# Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()
