from spire.presentation.common import *
from spire.presentation import *


outputFile = "CreateScatterChart.pptx"

#Creat a presentation
pres = Presentation()

#Set background image
ImageFile = "Data/bg.png"
rect2 = RectangleF.FromLTRB (0, 0, pres.SlideSize.Size.Width, pres.SlideSize.Size.Height)
pres.Slides[0].Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, ImageFile, rect2)
pres.Slides[0].Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.get_FloralWhite()

#Insert a chart and set chart title and chart type
rect1 = RectangleF.FromLTRB (90, 100, 640, 420)
chart = pres.Slides[0].Shapes.AppendChartInit (ChartType.ScatterMarkers, rect1, False)
chart.ChartTitle.TextProperties.Text = "ScatterMarker Chart"
chart.ChartTitle.TextProperties.IsCentered = True
chart.ChartTitle.Height = 30
chart.HasTitle = True

#Set chart data
xdata = [2.7, 8.9, 10.0, 12.4]
ydata = [3.2, 15.3, 6.7, 8]
chart.ChartData[0,0].Text = "X-Value"
chart.ChartData[0,1].Text = "Y-Value"
i = 0
while i < len(xdata):
    chart.ChartData[i + 1,0].NumberValue = xdata[i]
    chart.ChartData[i + 1,1].NumberValue = ydata[i]
    i += 1

#Set the series label
chart.Series.SeriesLabel = chart.ChartData["B1","B1"]

#Assign data to X axis, Y axis and Bubbles
chart.Series[0].XValues = chart.ChartData["A2","A5"]
chart.Series[0].YValues = chart.ChartData["B2","B5"]

pres.SaveToFile(outputFile, FileFormat.Pptx2010)
pres.Dispose()