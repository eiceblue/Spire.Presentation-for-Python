from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/TwoColumns.pptx"
outputFile ="SetRightToLeftColumns.pptx"

ppt = Presentation()
ppt.LoadFromFile(inputFile)
#Get the second shape
shape = ppt.Slides[0].Shapes[1] if isinstance(ppt.Slides[0].Shapes[1], IAutoShape) else None
#Set columns style to right-to-left
shape.TextFrame.RightToLeftColumns = True
#Save to file.
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()
