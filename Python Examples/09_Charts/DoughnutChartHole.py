from spire.presentation.common import *
from spire.presentation import *

inputFile = "Data/DoughnutChart.pptx"
outputFile = "DoughnutChartHole.pptx"

#Create PPT document and load file
ppt = Presentation()
ppt.LoadFromFile(inputFile)

#Get the chart on the first slide
Chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Set hole size
Chart.Series[0].DoughnutHoleSize = 55

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2010)
ppt.Dispose()
