from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/ChartSample2.pptx"
outputFile = "SetDatalabelPosition.pptx"

#Create PPT document and load file
ppt = Presentation()
ppt.LoadFromFile(inputFile)

#Get chart on the first slide
Chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Add data label
label = Chart.Series[0].DataLabels.Add()
#Set the position of the label
label.X = 0.1
label.Y = 0.1

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2010)
ppt.Dispose()

