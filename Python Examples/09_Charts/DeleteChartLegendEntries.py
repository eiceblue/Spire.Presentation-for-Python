from spire.presentation.common import *
from spire.presentation import *

inputFile = "./Data/Template_Ppt_2.pptx"
outputFile = "DeleteChartLegendEntries.pptx"

#Create a PowerPoint document.
presentation = Presentation()

#Load the file from disk.
presentation.LoadFromFile(inputFile)

#Get the chart.
chart = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IChart) else None

#Delete the first and the second legend entries from the chart.
chart.ChartLegend.DeleteEntry(0)
chart.ChartLegend.DeleteEntry(1)

#Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()