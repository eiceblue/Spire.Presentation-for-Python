from spire.presentation.common import *
from spire.presentation import *


inputFile = "Data/Template_Ppt_2.pptx"
outputFile = "AddTrendLineForChartSeries.pptx"

#Create a PPT document
presentation = Presentation()

#Load the file from disk.
presentation.LoadFromFile(inputFile)

#Get the target chart, add trendline for the first data series of the chart and specify the trendline type.
chart = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IChart) else None
it = chart.Series[0].AddTrendLine(TrendlinesType.Linear)

#Set the trendline properties to determine what should be displayed.
it.displayEquation = False
it.displayRSquaredValue = False

#Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()