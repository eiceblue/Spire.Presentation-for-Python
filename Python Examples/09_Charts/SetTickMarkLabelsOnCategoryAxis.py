from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/Template_Ppt_3.pptx"
outputFile = "Template_Ppt_3.pptx"

#Create a PowerPonit document.
presentation = Presentation()

#Load the file from disk.
presentation.LoadFromFile(inputFile)

#Get the chart from the PowerPoint slide.
chart = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IChart) else None

#Rotate tick labels.
chart.PrimaryCategoryAxis.TextRotationAngle = 45

#Specify interval between labels.
chart.PrimaryCategoryAxis.IsAutomaticTickLabelSpacing = False
chart.PrimaryCategoryAxis.TickLabelSpacing = 2

#Change position.
chart.PrimaryCategoryAxis.TickLabelPosition = TickLabelPositionType.TickLabelPositionHigh

#Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()

