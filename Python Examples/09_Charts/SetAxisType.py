from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/SetAxisType.pptx"
outputFile = "SetAxisType.pptx"

#Create an instance of presentation document
ppt = Presentation()
#Load file
ppt.LoadFromFile(inputFile)

#Get the chart
chart = ppt.Slides[0].Shapes[1] if isinstance(ppt.Slides[0].Shapes[1], IChart) else None

chart.PrimaryCategoryAxis.AxisType = AxisType.DateAxis
chart.PrimaryCategoryAxis.MajorUnitScale = ChartBaseUnitType.Months

#Save to file.
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()


