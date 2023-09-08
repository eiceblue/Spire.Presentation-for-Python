from spire.presentation.common import *
from spire.presentation import *


inputFile = "./Data/ExplodePieChart.pptx"
outputFile = "ExplodePieChart.pptx"

#Create a PowerPoint document.
presentation = Presentation()

#Load the file from disk.
presentation.LoadFromFile(inputFile)

#Get the chart that needs to set the point explosion.
chart = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IChart) else None

chart.Series[0].Distance = 15

#Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()