from spire.presentation.common import *
from spire.presentation import *


outputFile = "DoughnutChart.pptx"

#Create a ppt document
presentation = Presentation()
rect = RectangleF.FromLTRB (80, 100, 630, 420)

#Set background image
ImageFile = "Data/bg.png"
rect2 = RectangleF.FromLTRB (0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height)
presentation.Slides[0].Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, ImageFile, rect2)
presentation.Slides[0].Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.get_FloralWhite()

#Add a Doughnut chart
chart = presentation.Slides[0].Shapes.AppendChartInit(ChartType.Doughnut, rect, False)
chart.ChartTitle.TextProperties.Text = "Market share by country"
chart.ChartTitle.TextProperties.IsCentered = True
chart.ChartTitle.Height = 30
countries = ["Guba", "Mexico", "France", "German"]
sales = [1800, 3000, 5100, 6200]
chart.ChartData[0,0].Text = "Countries"
chart.ChartData[0,1].Text = "Sales"
i = 0
while i < len(countries):
    chart.ChartData[i + 1,0].Text = countries[i]
    chart.ChartData[i + 1,1].NumberValue = sales[i]
    i += 1
chart.Series.SeriesLabel = chart.ChartData["B1","B1"]
chart.Categories.CategoryLabels = chart.ChartData["A2","A5"]
chart.Series[0].Values = chart.ChartData["B2","B5"]
for i, item in enumerate(chart.Series[0].Values):
    cdp = ChartDataPoint(chart.Series[0])
    cdp.Index = i
    chart.Series[0].DataPoints.Add(cdp)

#Set the series color
chart.Series[0].DataPoints[0].Fill.FillType = FillFormatType.Solid
chart.Series[0].DataPoints[0].Fill.SolidColor.Color = Color.get_LightBlue()
chart.Series[0].DataPoints[1].Fill.FillType = FillFormatType.Solid
chart.Series[0].DataPoints[1].Fill.SolidColor.Color = Color.get_MediumPurple()
chart.Series[0].DataPoints[2].Fill.FillType = FillFormatType.Solid
chart.Series[0].DataPoints[2].Fill.SolidColor.Color = Color.get_DarkGray()
chart.Series[0].DataPoints[3].Fill.FillType = FillFormatType.Solid
chart.Series[0].DataPoints[3].Fill.SolidColor.Color = Color.get_DarkOrange()
chart.Series[0].DataLabels.LabelValueVisible = True
chart.Series[0].DataLabels.PercentValueVisible = True
chart.Series[0].DoughnutHoleSize = 60

presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()