from spire.presentation.common import *
from spire.presentation import *


inputFile = "Data/ChartSample2.pptx"
outputFile ="ChangeSeriesName.pptx"

#Create PPT document and load file
ppt = Presentation()
ppt.LoadFromFile(inputFile)

#Get chart on the first slide
Chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Get the ranges of series label 
cr = Chart.Series.SeriesLabel

#Change the value
cr[0].Text = "Changed series name"

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2010)
ppt.Dispose()
