from spire.presentation.common import *
from spire.presentation import *


inputFile = "./Data/Template_Ppt_2.pptx"
outputFile = "AddSecondaryValueAxisToChart.pptx"

#Create a PPT document
presentation = Presentation()

#Load the file from disk.
presentation.LoadFromFile(inputFile)

#Get the chart from the PowerPoint file.
chart = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IChart) else None

#Add a secondary axis to display the value of Series 3.
chart.Series[2].UseSecondAxis = True

#Set the grid line of secondary axis as invisible.
chart.SecondaryValueAxis.MajorGridTextLines.FillType = FillFormatType.none

#Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()