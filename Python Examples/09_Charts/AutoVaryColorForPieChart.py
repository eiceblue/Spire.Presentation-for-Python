from spire.presentation.common import *
from spire.presentation import *


outputFile = "AutoVaryColorForPieChart.pptx"

ppt = Presentation()
rect1 = RectangleF.FromLTRB (40, 100, 550+40, 320+100)

#Add a pie chart
chart = ppt.Slides[0].Shapes.AppendChartInit (ChartType.Pie, rect1, False)
chart.ChartTitle.TextProperties.Text = "Sales by Quarter"
chart.ChartTitle.TextProperties.IsCentered = True
chart.ChartTitle.Height = 30
chart.HasTitle = True

#Attach the data to chart
quarters = ["1st Qtr", "2nd Qtr", "3rd Qtr", "4th Qtr"]
sales = [210, 320, 180, 500]
chart.ChartData[0,0].Text = "Quarters"
chart.ChartData[0,1].Text = "Sales"
i = 0
while i < len(quarters):
    chart.ChartData[i + 1,0].Text = quarters[i]
    chart.ChartData[i + 1,1].NumberValue = sales[i]
    i += 1
chart.Series.SeriesLabel = chart.ChartData["B1","B1"]
chart.Categories.CategoryLabels = chart.ChartData["A2","A5"]
chart.Series[0].Values = chart.ChartData["B2","B5"]

#Set whether auto vary color, default value is true
chart.Series[0].IsVaryColor = False
chart.Series[0].Distance = 15

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2010)
ppt.Dispose()

