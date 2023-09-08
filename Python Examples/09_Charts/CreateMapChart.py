from spire.presentation.common import *
from spire.presentation import *


outputFile = "CreateMapChart.pptx"

#Create a PPT document
ppt = Presentation()

#Insert a Map chart to the first slide 
chart = ppt.Slides[0].Shapes.AppendChartInit (ChartType.Map, RectangleF.FromLTRB (50, 50, 500, 500), False)
chart.ChartData[0,1].Text = "series"

#Define some data.
countries = ["China", "Russia", "France", "Mexico", "United States", "India", "Australia"]
i = 0
while i < len(countries):
    chart.ChartData[i + 1,0].Text = countries[i]
    i += 1
values = [32, 20, 23, 17, 18, 6, 11]
i = 0
while i < len(values):
    chart.ChartData[i + 1,1].NumberValue = values[i]
    i += 1
chart.Series.SeriesLabel = chart.ChartData[0,1,0,1]
chart.Categories.CategoryLabels = chart.ChartData[1,0,7,0]
chart.Series[0].Values = chart.ChartData[1,1,7,1]

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()