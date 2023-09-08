from spire.presentation.common import *
from spire.presentation import *


inputFile = "Data/GroupTwoLevelAxisLabels.pptx"
outputFile = "GroupTwoLevelAxisLabels.pptx"

#Create a PowerPoint document.
presentation = Presentation()

#Load the file from disk.
presentation.LoadFromFile(inputFile)

#Get the chart.
chart = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IChart) else None

#Get the category axis from the chart.
chartAxis = chart.PrimaryCategoryAxis

#Group the axis labels that have the same first-level label.
if chartAxis.HasMultiLvlLbl:
    chartAxis.IsMergeSameLabel = True

#Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()

