from spire.presentation.common import *
from spire.presentation import *


outputFile = "CreatePieChart.pptx"

#Create a PPT document
presentation = Presentation()

#Insert a Pie chart to the first slide and set the chart title.
rect1 = RectangleF.FromLTRB (40, 100, 590, 420)
chart = presentation.Slides[0].Shapes.AppendChartInit (ChartType.Pie, rect1, False)
chart.ChartTitle.TextProperties.Text = "Sales by Quarter"
chart.ChartTitle.TextProperties.IsCentered = True
chart.ChartTitle.Height = 30
chart.HasTitle = True

#Define some data.
quarters = ["1st Qtr", "2nd Qtr", "3rd Qtr", "4th Qtr"]
sales = [210, 320, 180, 500]

#Append data to ChartData, which represents a data table where the chart data is stored.
chart.ChartData[0,0].Text = "Quarters"
chart.ChartData[0,1].Text = "Sales"
i = 0
while i < len(quarters):
    chart.ChartData[i + 1,0].Text = quarters[i]
    chart.ChartData[i + 1,1].NumberValue = sales[i]
    i += 1

#Set category labels, series label and series data.
chart.Series.SeriesLabel = chart.ChartData["B1","B1"]
chart.Categories.CategoryLabels = chart.ChartData["A2","A5"]
chart.Series[0].Values = chart.ChartData["B2","B5"]

#Add data points to series and fill each data point with different color.
for i, unusedItem in enumerate(chart.Series[0].Values):
    cdp = ChartDataPoint(chart.Series[0])
    cdp.Index = i
    chart.Series[0].DataPoints.Add(cdp)
chart.Series[0].DataPoints[0].Fill.FillType = FillFormatType.Solid
chart.Series[0].DataPoints[0].Fill.SolidColor.Color = Color.get_RosyBrown()
chart.Series[0].DataPoints[1].Fill.FillType = FillFormatType.Solid
chart.Series[0].DataPoints[1].Fill.SolidColor.Color = Color.get_LightBlue()
chart.Series[0].DataPoints[2].Fill.FillType = FillFormatType.Solid
chart.Series[0].DataPoints[2].Fill.SolidColor.Color = Color.get_LightPink()
chart.Series[0].DataPoints[3].Fill.FillType = FillFormatType.Solid
chart.Series[0].DataPoints[3].Fill.SolidColor.Color = Color.get_MediumPurple()

#Set the data labels to display label value and percentage value.
chart.Series[0].DataLabels.LabelValueVisible = True
chart.Series[0].DataLabels.PercentValueVisible = True

#Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()