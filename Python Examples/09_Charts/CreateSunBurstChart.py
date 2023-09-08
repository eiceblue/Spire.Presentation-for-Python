from spire.presentation.common import *
from spire.presentation import *


outputFile = "CreateSunBurstChart.pptx"

#Create PPT document
ppt = Presentation()

#Create a SunBurst chart to the first slide
chart = ppt.Slides[0].Shapes.AppendChartInit (ChartType.SunBurst, RectangleF.FromLTRB (50, 50, 550, 450), False)

#Set series text
chart.ChartData[0,3].Text = "Series 1"

#Set category text
categories = [["Branch 1", "Stem 1", "Leaf 1"], ["Branch 1", "Stem 1", "Leaf 2"], ["Branch 1", "Stem 1", "Leaf 3"], ["Branch 1", "Stem 2", "Leaf 4"], ["Branch 1", "Stem 2", "Leaf 5"], ["Branch 1", "Leaf 6", None], ["Branch 1", "Leaf 7", None], ["Branch 2", "Stem 3", "Leaf 8"], ["Branch 2", "Leaf 9", None], ["Branch 2", "Stem 4", "Leaf 10"], ["Branch 2", "Stem 4", "Leaf 11"], ["Branch 2", "Stem 5", "Leaf 12"], ["Branch 3", "Stem 5", "Leaf 13"], ["Branch 3", "Stem 6", "Leaf 14"], ["Branch 3", "Leaf 15", None]]
for i in range(0, 15):
    for j in range(0, 3):
        chart.ChartData[i + 1,j].Text = categories[i][j]

#Fill data for chart
values = [17, 23, 48, 22, 76, 54, 77, 26, 44, 63, 10, 15, 48, 15, 51]
i = 0
while i < len(values):
    chart.ChartData[i + 1,3].NumberValue = values[i]
    i += 1

#Set series labels
chart.Series.SeriesLabel = chart.ChartData[0,3,0,3]

#Set categories labels 
chart.Categories.CategoryLabels = chart.ChartData[1,0,len(values),2]

#Assign data to series values
chart.Series[0].Values = chart.ChartData[1,3,len(values),3]
chart.Series[0].DataLabels.CategoryNameVisible = True
chart.ChartTitle.TextProperties.Text = "SunBurst"
chart.HasLegend = True
chart.ChartLegend.Position = ChartLegendPositionType.Top

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()
