from spire.presentation.common import *
from spire.presentation import *


outputFile = "CreateClusteredColumnChart.pptx"

#Create a PPT file
presentation = Presentation()

#Add clustered column chart
rect1 = RectangleF.FromLTRB (90, 100, 640, 420)
chart = presentation.Slides[0].Shapes.AppendChartInit (ChartType.ColumnClustered, rect1, False)

#Chart title
chart.ChartTitle.TextProperties.Text = "Clustered Column Chart"
chart.ChartTitle.TextProperties.IsCentered = True
chart.ChartTitle.Height = 30
chart.HasTitle = True

#Data for series
Series1 = [7.7, 8.9, 1.0, 2.4]
Series2 = [15.2, 5.3, 6.7, 8]

#Set series text
chart.ChartData[0,1].Text = "Series1"
chart.ChartData[0,2].Text = "Series2"

#Set category text
chart.ChartData[1,0].Text = "Category 1"
chart.ChartData[2,0].Text = "Category 2"
chart.ChartData[3,0].Text = "Category 3"
chart.ChartData[4,0].Text = "Category 4"

#Fill data for chart
i = 0
while i < len(Series1):
    chart.ChartData[i + 1,1].NumberValue = Series1[i]
    chart.ChartData[i + 1,2].NumberValue = Series2[i]
    i += 1

#Set series label
chart.Series.SeriesLabel = chart.ChartData["B1","C1"]

#Set category label
chart.Categories.CategoryLabels = chart.ChartData["A2","A5"]

#Set values for series
chart.Series[0].Values = chart.ChartData["B2","B5"]
chart.Series[1].Values = chart.ChartData["C2","C5"]

#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()
