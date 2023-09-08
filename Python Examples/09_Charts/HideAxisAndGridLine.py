from spire.presentation.common import *
from spire.presentation import *

inputFile ="Data/ChartSample2.pptx"
outputFile ="HideAxisAndGridLine.pptx"

#Create PPT document and load file
ppt = Presentation()
ppt.LoadFromFile(inputFile)

#Get chart on the first slide
Chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Hide axis
Chart.PrimaryCategoryAxis.IsVisible = False
Chart.PrimaryValueAxis.IsVisible = False

#Remove gridline
Chart.PrimaryValueAxis.MajorGridTextLines.FillType = FillFormatType.none

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2010)
ppt.Dispose()