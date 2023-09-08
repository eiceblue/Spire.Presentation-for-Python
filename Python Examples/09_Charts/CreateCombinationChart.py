from spire.presentation.common import *
from spire.presentation import *


outputFile ="CombinationChart.pptx"

#Create a presentation instance
presentation = Presentation()

#Set background image
ImageFile = "Data/bg.png"
rect2 = RectangleF.FromLTRB (0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height)
presentation.Slides[0].Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, ImageFile, rect2)
presentation.Slides[0].Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.get_FloralWhite()

#Insert a column clustered chart
rect = RectangleF.FromLTRB (100, 100, 650, 420)
chart = presentation.Slides[0].Shapes.AppendChart(ChartType.ColumnClustered, rect)

#Set chart title
chart.ChartTitle.TextProperties.Text = "Monthly Sales Report"
chart.ChartTitle.TextProperties.IsCentered = True
chart.ChartTitle.Height = 30
chart.HasTitle = True

caption = ["Month","Sales","Growth rate"]
month =["January","February","March","April","May","June"]
sales =[200,250,300,150,200,400]
growth_rate = [0.6,0.8,0.6,0.2,0.5,0.9]

#Import data from datatable to chart data
for c,text in enumerate(caption):
    chart.ChartData[0,c].Text = text
for m,text in enumerate(month):
    chart.ChartData[m+1,0].Text = text
for s,num in enumerate(sales):
    chart.ChartData[s+1,1].NumberValue = num
for g,num in enumerate(growth_rate):
    chart.ChartData[g+1,2].NumberValue = num

#Set series labels
chart.Series.SeriesLabel = chart.ChartData["B1","C1"]

#Set categories labels    
chart.Categories.CategoryLabels = chart.ChartData["A2","A7"]

#Assign data to series values
chart.Series[0].Values = chart.ChartData["B2","B7"]
chart.Series[1].Values = chart.ChartData["C2","C7"]

#Change the chart type of serie 2 to line with markers
chart.Series[1].Type = ChartType.LineMarkers

#Plot data of series 2 on the secondary axis
chart.Series[1].UseSecondAxis = True

#Set the number format as percentage 
chart.SecondaryValueAxis.NumberFormat = "0%"

#Hide gridlinkes of secondary axis
chart.SecondaryValueAxis.MajorGridTextLines.FillType = FillFormatType.none

#Set overlap
chart.OverLap = -50

#Set gapwidth
chart.GapWidth = 200

#Save to file
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()