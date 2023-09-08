from spire.presentation.common import *
from spire.presentation import *

inputFile = "./Data/LastViewedSlide.pptx"
outputFile = "LinkToLastViewedSlide.pptx"

ppt = Presentation()
ppt.LoadFromFile(inputFile)

slide = ppt.Slides[0]
# Draw a shape
autoShape = slide.Shapes.AppendShape(
    ShapeType.Rectangle, RectangleF.FromLTRB(100, 100, 200, 200))
# Link to last viewed slide show
autoShape.Click = ClickHyperlink.get_LastVievedSlide()

# Save to file.
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()
