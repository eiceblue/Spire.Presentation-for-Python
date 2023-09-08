from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/ScaleBubbleChartSize.pptx"
outputFile = "ScaleBubbleChartSize.pptx"

#Create a PowerPoint document.
presentation = Presentation()

#Load the file from disk.
presentation.LoadFromFile(inputFile)

#Get the chart from the first presentation slide.
chart = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IChart) else None

#Scale the bubble size, the range value is from 0 to 300.
chart.BubbleScale = 50

#Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()

