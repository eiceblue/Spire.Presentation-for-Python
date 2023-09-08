from spire.presentation.common import *
from spire.presentation import *


outputFile = "CreateWaterFallChart.pptx"

#Create PPT document
ppt = Presentation()

#Create a WaterFall chart to the first slide
chart = ppt.Slides[0].Shapes.AppendChartInit (ChartType.WaterFall, RectangleF.FromLTRB (50, 50, 550, 450), False)

#Set series text
chart.ChartData[0,1].Text = "Series 1"

#Set category text
categories = ["Category 1", "Category 2", "Category 3", "Category 4", "Category 5", "Category 6", "Category 7"]
i = 0
while i < len(categories):
    chart.ChartData[i + 1,0].Text = categories[i]
    i += 1

#Fill data for chart
values = [100, 20, 50, -40, 130, -60, 70]
i = 0
while i < len(values):
    chart.ChartData[i + 1,1].NumberValue = values[i]
    i += 1

#Set series labels
chart.Series.SeriesLabel = chart.ChartData[0,1,0,1]

#Set categories labels 
chart.Categories.CategoryLabels = chart.ChartData[1,0,len(categories),0]

#Assign data to series values
chart.Series[0].Values = chart.ChartData[1,1,len(values),1]

#Operate the third datapoint of first series
chartDataPoint = ChartDataPoint(chart.Series[0])
chartDataPoint.Index = 2
chartDataPoint.SetAsTotal = True
chart.Series[0].DataPoints.Add(chartDataPoint)

#Operate the sixth datapoint of first series
chartDataPoint2 = ChartDataPoint(chart.Series[0])
chartDataPoint2.Index = 5
chartDataPoint2.SetAsTotal = True
chart.Series[0].DataPoints.Add(chartDataPoint2)
chart.Series[0].ShowConnectorLines = True
chart.Series[0].DataLabels.LabelValueVisible = True
chart.ChartLegend.Position = ChartLegendPositionType.Right
chart.ChartTitle.TextProperties.Text = "WaterFall"

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()
