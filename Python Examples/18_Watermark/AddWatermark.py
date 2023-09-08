from spire.presentation.common import *
from spire.presentation import *

inputFile = "Data/AddWatermark.pptx"
outputFile = "output/Watermark.pptx"

# Create a PPT document and load file
presentation = Presentation()
presentation.LoadFromFile(inputFile)

# Define a rectangle range
left = (presentation.SlideSize.Size.Width - 336.4) / 2
top = (presentation.SlideSize.Size.Height - 110.8) / 2
rect = RectangleF(left, top, 336.4, 110.8)

# Add a rectangle shape with a defined range
shape = presentation.Slides[0].Shapes.AppendShape(
    ShapeType.Rectangle, rect)

# Set the style of the shape
shape.Fill.FillType = FillFormatType.none
shape.ShapeStyle.LineColor.Color = Color.get_White()
shape.Rotation = -45
shape.Locking.SelectionProtection = True
shape.Line.FillType = FillFormatType.none

# Add text to the shape
shape.TextFrame.Text = "E-iceblue"
textRange = shape.TextFrame.TextRange
# Set the style of the text range
textRange.Fill.FillType = FillFormatType.Solid
textRange.Fill.SolidColor.Color = Color.FromArgb(
    120, Color.get_HotPink().R, Color.get_HotPink().G, Color.get_HotPink().B)
textRange.FontHeight = 50

# Save the document and launch
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()
