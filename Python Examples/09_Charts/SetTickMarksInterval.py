from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/SetTickMarksInterval.pptx"
outputFile = "SetTickMarksInterval.pptx"

#Create PPT document
ppt = Presentation()
#Load file
ppt.LoadFromFile(inputFile)
chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None
chartAxis = chart.PrimaryCategoryAxis
chartAxis.TickMarkSpacing = 2

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()

