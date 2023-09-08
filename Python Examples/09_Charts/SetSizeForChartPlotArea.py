from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/ChartSample2.pptx"
outputFile = "SetSizeForChartPlotArea.pptx"

#Create PPT document and load file
ppt = Presentation()
ppt.LoadFromFile(inputFile)

#Get chart on the first slide
Chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Set width and height for chart plot area
Chart.PlotArea.Width = 250
Chart.PlotArea.Height = 300

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2010)
ppt.Dispose()

