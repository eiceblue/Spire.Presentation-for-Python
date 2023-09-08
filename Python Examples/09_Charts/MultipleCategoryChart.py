from spire.presentation.common import *
from spire.presentation import *


outputFile ="MultipleCategoryChart.pptx"

#Create a PPT file
presentation = Presentation()

#Add line markers chart
rect1 = RectangleF.FromLTRB (90, 100, 640, 420)
chart = presentation.Slides[0].Shapes.AppendChartInit (ChartType.ColumnClustered, rect1, False)

#Chart title
chart.ChartTitle.TextProperties.Text = "Muli-Category"
chart.ChartTitle.TextProperties.IsCentered = True
chart.ChartTitle.Height = 30
chart.HasTitle = True


#Data for series
Series1 = [7.7, 8.9, 7, 6, 7, 8]

#Set series text
chart.ChartData[0,2].Text = "Series1"

#Set category text
chart.ChartData[1,0].Text = "Grp 1"
chart.ChartData[3,0].Text = "Grp 2"
chart.ChartData[5,0].Text = "Grp 3"

chart.ChartData[1,1].Text = "A"
chart.ChartData[2,1].Text = "B"
chart.ChartData[3,1].Text = "C"
chart.ChartData[4,1].Text = "D"
chart.ChartData[5,1].Text = "E"
chart.ChartData[6,1].Text = "F"


#Fill data for chart
i = 0
while i < len(Series1):
    chart.ChartData[i + 1,2].NumberValue = Series1[i]
    i += 1

#Set series label
chart.Series.SeriesLabel = chart.ChartData["C1","C1"]
#Set category label
chart.Categories.CategoryLabels = chart.ChartData["A2","B7"]

#Set values for series
chart.Series[0].Values = chart.ChartData["C2","C7"]

#Set if the category axis has multiple levels
chart.PrimaryCategoryAxis.HasMultiLvlLbl = True
#Merge same label
chart.PrimaryCategoryAxis.IsMergeSameLabel = True

#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()