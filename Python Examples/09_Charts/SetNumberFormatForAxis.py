from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/ChartSample3.pptx"
outputFile = "SetNumberFormatForAxis.pptx"

#Create PPT document and load file
ppt = Presentation()
ppt.LoadFromFile(inputFile)

#Get chart on the first slide
Chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Set the number format
Chart.PrimaryCategoryAxis.NumberFormat = "yyyy"

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2010)
ppt.Dispose()
