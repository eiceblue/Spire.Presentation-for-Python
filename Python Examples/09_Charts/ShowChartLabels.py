from spire.presentation.common import *
from spire.presentation import *


inputFile ="./Data/ChartSample2.pptx"
outputFile = "ShowChartLabels.pptx"

#Create PPT document and load file
ppt = Presentation()
ppt.LoadFromFile(inputFile)

#Get chart on the first slide
Chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Show data labels
Chart.Series[0].DataLabels.LabelValueVisible = True
Chart.Series[0].DataLabels.CategoryNameVisible = True
Chart.Series[0].DataLabels.SeriesNameVisible = True

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2010)
ppt.Dispose()

