from spire.presentation.common import *
from spire.presentation import *


outputFile = "CreateParetoChart.pptx"

#Create PPT document
ppt = Presentation()

#Create a Pareto chart in first slide
chart = ppt.Slides[0].Shapes.AppendChartInit (ChartType.Pareto, RectangleF.FromLTRB (50, 50, 550, 450), False)

#Set series text
chart.ChartData[0,1].Text = "Series 1"

#Set category text
categories = ["Category 1", "Category 2", "Category 4", "Category 3", "Category 4", "Category 2", "Category 1", "Category 1", "Category 3", "Category 2", "Category 4", "Category 2", "Category 3", "Category 1", "Category 3", "Category 2", "Category 4", "Category 1", "Category 1", "Category 3", "Category 2", "Category 4", "Category 1", "Category 1", "Category 3", "Category 2", "Category 4", "Category 1"]
i = 0
while i < len(categories):
    chart.ChartData[i + 1,0].Text = categories[i]
    i += 1

#Fill data for chart
values = [1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1]
i = 0
while i < len(values):
    chart.ChartData[i + 1,1].NumberValue = values[i]
    i += 1
chart.Series.SeriesLabel = chart.ChartData[0,1,0,1]
chart.Categories.CategoryLabels = chart.ChartData[1,0,len(categories),0]
chart.Series[0].Values = chart.ChartData[1,1,len(values),1]
chart.PrimaryCategoryAxis.IsBinningByCategory = True
chart.Series[1].Line.FillFormat.FillType = FillFormatType.Solid
chart.Series[1].Line.FillFormat.SolidFillColor.Color = Color.get_Red()
chart.ChartTitle.TextProperties.Text = "Pareto"
chart.HasLegend = True
chart.ChartLegend.Position = ChartLegendPositionType.Bottom

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()