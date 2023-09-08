from spire.presentation.common import *
from spire.presentation import *


outputFile = "LinkToASpecificSlide.pptx"

# Create a PowerPoint document.
presentation = Presentation()

# Append a slide to it.
presentation.Slides.Append()

# Add a shape to the second slide.
shape = presentation.Slides[1].Shapes.AppendShape(
    ShapeType.Rectangle, RectangleF.FromLTRB(10, 50, 210, 100))
shape.Fill.FillType = FillFormatType.none
shape.Line.FillType = FillFormatType.none
shape.TextFrame.Text = "Jump to the first slide"

# Create a hyperlink based on the shape and the text on it, linking to the first slide.
hyperlink = ClickHyperlink(presentation.Slides[0])
shape.Click = hyperlink
shape.TextFrame.TextRange.ClickAction = hyperlink

# Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()
