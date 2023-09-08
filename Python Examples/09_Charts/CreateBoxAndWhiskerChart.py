from spire.presentation.common import *
from spire.presentation import *


outputFile = "CreateBoxAndWhiskerChart.pptx"

# Create a PPT document
ppt = Presentation()

# Insert a BoxAndWhisker chart to the first slide 
chart = ppt.Slides[0].Shapes.AppendChartInit(ChartType.BoxAndWhisker, RectangleF.FromLTRB(50, 50, 550, 450), False)

# Series labels
seriesLabel = ["Series 1", "Series 2", "Series 3"]
i = 0
while i < len(seriesLabel):
    chart.ChartData[0,i + 1].Text = "Series 1"
    i += 1

# Categories
categories = ["Category 1", "Category 1", "Category 1", "Category 1", "Category 1", "Category 1", "Category 1", "Category 2", "Category 2", "Category 2", "Category 2", "Category 2", "Category 2", "Category 3", "Category 3", "Category 3", "Category 3", "Category 3"]
i = 0
while i < len(categories):
    chart.ChartData[i + 1,0].Text = categories[i]
    i += 1

# Values
values = [[-7, -3, -24], [-10, 1, 11], [-28, -6, 34], [47, 2, -21], [35, 17, 22], [-22, 15, 19], [17, -11, 25], [-30, 18, 25], [49, 22, 56], [37, 22, 15], [-55, 25, 31], [14, 18, 22], [18, -22, 36], [-45, 25, -17], [-33, 18, 22], [18, 2, -23], [-33, -22, 10], [10, 19, 22]]
i = 0
while i < len(seriesLabel):
    j = 0
    while j < len(categories):
        chart.ChartData[j + 1,i + 1].NumberValue = values[j][i]
        j += 1
    i += 1
chart.Series.SeriesLabel = chart.ChartData[0,1,0,len(seriesLabel)]
chart.Categories.CategoryLabels = chart.ChartData[1,0,len(categories),0]
chart.Series[0].Values = chart.ChartData[1,1,len(categories),1]
chart.Series[1].Values = chart.ChartData[1,2,len(categories),2]
chart.Series[2].Values = chart.ChartData[1,3,len(categories),3]
chart.Series[0].ShowInnerPoints = False
chart.Series[0].ShowOutlierPoints = True
chart.Series[0].ShowMeanMarkers = True
chart.Series[0].ShowMeanLine = True
chart.Series[0].QuartileCalculationType = QuartileCalculation.ExclusiveMedian
chart.Series[1].ShowInnerPoints = False
chart.Series[1].ShowOutlierPoints = True
chart.Series[1].ShowMeanMarkers = True
chart.Series[1].ShowMeanLine = True
chart.Series[1].QuartileCalculationType = QuartileCalculation.InclusiveMedian
chart.Series[2].ShowInnerPoints = False
chart.Series[2].ShowOutlierPoints = True
chart.Series[2].ShowMeanMarkers = True
chart.Series[2].ShowMeanLine = True
chart.Series[2].QuartileCalculationType = QuartileCalculation.ExclusiveMedian
chart.HasLegend = True
chart.ChartTitle.TextProperties.Text = "BoxAndWhisker"
chart.ChartLegend.Position = ChartLegendPositionType.Top

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()
