from spire.presentation.common import *
from spire.presentation import *


outputFile = "CreateHistogramChart.pptx"

#Create PPT document
ppt = Presentation()

#Add a Histogram chart
chart = ppt.Slides[0].Shapes.AppendChartInit (ChartType.Histogram, RectangleF.FromLTRB (50, 50, 550, 450), False)

#Set series text
chart.ChartData[0,0].Text = "Series 1"

#Fill data for chart
values = [1, 1, 1, 3, 3, 3, 3, 5, 5, 5, 8, 8, 8, 9, 9, 9, 12, 12, 13, 13, 17, 17, 17, 19, 19, 19, 25, 25, 25, 25, 25, 25, 25, 25, 29, 29, 29, 29, 32, 32, 33, 33, 35, 35, 41, 41, 44, 45, 49, 49]
i = 0
while i < len(values):
    chart.ChartData[i + 1,1].NumberValue = values[i]
    i += 1

#Set series label
chart.Series.SeriesLabel = chart.ChartData[0,0,0,0]

#Set values for series
chart.Series[0].Values = chart.ChartData[1,0,len(values),0]
chart.PrimaryCategoryAxis.NumberOfBins = 7
chart.PrimaryCategoryAxis.GapWidth = 20

#Chart title
chart.ChartTitle.TextProperties.Text = "Histogram"
chart.ChartLegend.Position = ChartLegendPositionType.Bottom

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()