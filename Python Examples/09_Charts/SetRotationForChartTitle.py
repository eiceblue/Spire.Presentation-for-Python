from spire.presentation.common import *
from spire.presentation import *


inputFile ="./Data/ChartSample2.pptx"
outputFile = "SetRotationForChartTitle.pptx"

#Create an instance of presentation document
ppt = Presentation()
#Load file
ppt.LoadFromFile(inputFile)

#Get the chart
chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

chart.ChartTitle.TextProperties.RotationAngle = -30

#Save to file.
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()



