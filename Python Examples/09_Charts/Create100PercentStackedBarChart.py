from spire.presentation.common import *
from spire.presentation import *


outputFile ="Create100PercentStackedBarChart.pptx"

#Create a PowerPoint document.
presentation = Presentation()

#Add a "Bar100PercentStacked" chart to the first slide.
presentation.SlideSize.Type = SlideSizeType.Screen16x9
slidesize = presentation.SlideSize.Size
slide = presentation.Slides[0]

#Append a chart.
rect = RectangleF.FromLTRB (20, 20, slidesize.Width - 20, slidesize.Height - 20)
chart = slide.Shapes.AppendChart(ChartType.Bar100PercentStacked, rect)

#Write data to the chart data.
columnlabels = ["Series 1", "Series 2", "Series 3"]

#Insert the column labels.
c = 0
while c < len(columnlabels):
    chart.ChartData[0,c + 1].Text = columnlabels[c]
    c += 1
rowlabels = ["Category 1", "Category 2", "Category 3"]

#Insert the row labels.
r = 0
while r < len(rowlabels):
    chart.ChartData[r + 1,0].Text = rowlabels[r]
    r += 1
values = [[ 20.83233, 10.34323, -10.354667 ], [ 10.23456, -12.23456, 23.34456 ], [ 12.34345, -23.34343, -13.23232 ]]

#Insert the values.
value = 0.0
r = 0
while r < len(rowlabels):
    c = 0
    while c < len(columnlabels):
        value = round(values[r][c], 2)
        chart.ChartData[r + 1,c + 1].NumberValue = value
        c += 1
    r += 1
chart.Series.SeriesLabel = chart.ChartData[0,1,0,len(columnlabels)]
chart.Categories.CategoryLabels = chart.ChartData[1,0,len(rowlabels),0]

#Set the position of category axis.
chart.PrimaryCategoryAxis.Position = AxisPositionType.Left
chart.SecondaryCategoryAxis.Position = AxisPositionType.Left
chart.PrimaryCategoryAxis.TickLabelPosition = TickLabelPositionType.TickLabelPositionLow

#Set the data, font and format for the series of each column.
c = 0
while c < len(columnlabels):
    chart.Series[c].Values = chart.ChartData[1,c + 1,len(rowlabels),c + 1]
    chart.Series[c].Fill.FillType = FillFormatType.Solid
    chart.Series[c].InvertIfNegative = False
    r = 0
    while r < len(rowlabels):
        label = chart.Series[c].DataLabels.Add()
        label.LabelValueVisible = True
        chart.Series[c].DataLabels[r].HasDataSource = False
        chart.Series[c].DataLabels[r].NumberFormat = "0#\\%"
        chart.Series[c].DataLabels.TextProperties.Paragraphs[0].DefaultCharacterProperties.FontHeight = 12
        r += 1
    c += 1

#Set the color of the Series.
chart.Series[0].Fill.SolidColor.Color = Color.get_YellowGreen()
chart.Series[1].Fill.SolidColor.Color = Color.get_Red()
chart.Series[2].Fill.SolidColor.Color = Color.get_Green()
font = TextFont("Tw Cen MT")

#Set the font and size for chartlegend.
k = 0
while k < len(chart.ChartLegend.EntryTextProperties):
    chart.ChartLegend.EntryTextProperties[k].LatinFont = font
    chart.ChartLegend.EntryTextProperties[k].FontHeight = 20
    k += 1
    
#Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()