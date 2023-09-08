from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/Template_Ppt_2.pptx"
outputFile = "SetNumberFormatAndRemoveTickMarksOfChart.pptx"

#Create a PowerPoint document.
presentation = Presentation()

#Load the file from disk.
presentation.LoadFromFile(inputFile)

#Get the chart that need to be adjusted the number format and remove the tick marks.
chart = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IChart) else None

#Set percentage number format for the axis value of chart.
chart.PrimaryValueAxis.NumberFormat = "0#\\%"

#Remove the tick marks for value axis and category axis.
chart.PrimaryValueAxis.MajorTickMark = TickMarkType.TickMarkNone
chart.PrimaryValueAxis.MinorTickMark = TickMarkType.TickMarkNone
chart.PrimaryCategoryAxis.MajorTickMark = TickMarkType.TickMarkNone
chart.PrimaryCategoryAxis.MinorTickMark = TickMarkType.TickMarkNone

#Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()


