from spire.presentation.common import *
from spire.presentation import *

inputFile = "Data/ChartSample2.pptx"
outputFile = "EditChartData.pptx"

#Create PPT document and load file
ppt = Presentation()
ppt.LoadFromFile(inputFile)

#Get chart on the first slide
Chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Change the value of the second datapoint of the first series
Chart.Series[0].Values[1].NumberValue = 6

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2010)
ppt.Dispose()

