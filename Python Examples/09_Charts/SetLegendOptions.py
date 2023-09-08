from spire.presentation.common import *
from spire.presentation import *


inputFile ="./Data/ChartSample2.pptx"
outputFile = "SetLegendOptions.pptx"

#Create PPT document and load file
ppt = Presentation()
ppt.LoadFromFile(inputFile)

#Get chart on the first slide
Chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Set the legend positon
Chart.ChartLegend.Left = 20
Chart.ChartLegend.Top = 20

#Set the legend size
Chart.ChartLegend.Width = 250
Chart.ChartLegend.Height = 30

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2010)
ppt.Dispose()

