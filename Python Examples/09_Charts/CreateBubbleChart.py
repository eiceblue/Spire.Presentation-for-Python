from spire.presentation.common import *
from spire.presentation import *


outputFile = "BubbleChart.pptx"

#Create a PPT file.
presentation = Presentation()

#Set background image
ImageFile = "Data/bg.png"
rect2 = RectangleF.FromLTRB (0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height)
presentation.Slides[0].Shapes.AppendEmbedImageByPath(ShapeType.Rectangle, ImageFile, rect2)
presentation.Slides[0].Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.get_FloralWhite()

#Add bubble chart
rect1 = RectangleF.FromLTRB (90, 100, 640, 420)
chart = presentation.Slides[0].Shapes.AppendChartInit(ChartType.Bubble, rect1, False)

#Chart title
chart.ChartTitle.TextProperties.Text = "Bubble Chart"
chart.ChartTitle.TextProperties.IsCentered = True
chart.ChartTitle.Height = 30
chart.HasTitle = True

#Attach the data to chart
xdata = [7.7, 8.9, 1.0, 2.4]
ydata = [15.2, 5.3, 6.7, 8]
size = [1.1, 2.4, 3.7, 4.8]
chart.ChartData[0,0].Text = "X-Value"
chart.ChartData[0,1].Text = "Y-Value"
chart.ChartData[0,2].Text = "Size"
i = 0
while i < len(xdata):
    chart.ChartData[i + 1,0].NumberValue = xdata[i]
    chart.ChartData[i + 1,1].NumberValue = ydata[i]
    chart.ChartData[i + 1,2].NumberValue = size[i]
    i += 1

#Set series label
chart.Series.SeriesLabel = chart.ChartData["B1","B1"]
chart.Series[0].XValues = chart.ChartData["A2","A5"]
chart.Series[0].YValues = chart.ChartData["B2","B5"]
chart.Series[0].Bubbles.Add(chart.ChartData["C2"])
chart.Series[0].Bubbles.Add(chart.ChartData["C3"])
chart.Series[0].Bubbles.Add(chart.ChartData["C4"])
chart.Series[0].Bubbles.Add(chart.ChartData["C5"])

presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()
