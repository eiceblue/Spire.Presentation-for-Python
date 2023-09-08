from spire.presentation.common import *
from spire.presentation import *


inputFile = "Data/ColumnChart.pptx"
outputFile = "InvertIfNegative.pptx"

#Create PPT document and load file
ppt = Presentation()
ppt.LoadFromFile(inputFile)

#Get chart on the first slide
Chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Set invert if negative
Chart.Series[0].InvertIfNegative = True

#Chart.Series[0].DataPoints[0].InvertIfNegative = true

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2010)
ppt.Dispose()