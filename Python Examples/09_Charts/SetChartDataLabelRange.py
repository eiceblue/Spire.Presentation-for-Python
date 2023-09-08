from spire.presentation.common import *
from spire.presentation import *

outputFile = "SetChartDataLabelRange.pptx"
#Create a PowerPoint document.
presentation = Presentation()

#Add a ColumnStacked chart
chart = presentation.Slides[0].Shapes.AppendChart(ChartType.ColumnStacked, RectangleF.FromLTRB (100, 100, 600, 500))

#Set data for the chart
cellRange = chart.ChartData["F1"]
cellRange.Text = "labelA"
cellRange = chart.ChartData["F2"]
cellRange.Text = "labelB"
cellRange = chart.ChartData["F3"]
cellRange.Text = "labelC"
cellRange = chart.ChartData["F4"]
cellRange.Text = "labelD"

#Set data label ranges
chart.Series[0].DataLabelRanges = chart.ChartData["F1","F4"]

#Add data label
dataLabel1 = chart.Series[0].DataLabels.Add()
dataLabel1.ID = 0
#Show the value
dataLabel1.LabelValueVisible = False
#Show the label string
dataLabel1.ShowDataLabelsRange = True

#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()

