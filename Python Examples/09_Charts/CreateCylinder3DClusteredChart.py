from spire.presentation.common import *
import math
from spire.presentation import *
import xml.etree.ElementTree as ET


outputFile = "CreateCylinder3DClusteredChart.pptx"

#Create a PPT document
presentation = Presentation()

#Set background image
ImageFile = "Data/bg.png"
rect2 = RectangleF.FromLTRB (0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height)
presentation.Slides[0].Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, ImageFile, rect2)
presentation.Slides[0].Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.get_FloralWhite()

#Insert chart
left = math.trunc(presentation.SlideSize.Size.Width / float(2)) - 200
rect = RectangleF.FromLTRB (left, 85, 400+left, 485)
chart = presentation.Slides[0].Shapes.AppendChart(ChartType.Cylinder3DClustered, rect)

#Add chart Title
chart.ChartTitle.TextProperties.Text = "Report"
chart.ChartTitle.TextProperties.IsCentered = True
chart.ChartTitle.Height = 30
chart.HasTitle = True

#Load data from XML file
caption =["SalesPers","SaleAmt","ComPct","ComAmt"]
for indx, cp in enumerate(caption):
    chart.ChartData[0,indx].Text = cp
tree = ET.parse("Data/data.xml")
root = tree.getroot()
for indx,child in enumerate(root):
    for i,subChild in enumerate(child):
        if(i==0):
            chart.ChartData[indx+1,i].Text = subChild.text
        else:
            chart.ChartData[indx+1,i].NumberValue = float(subChild.text)

chart.Series.SeriesLabel = chart.ChartData["B1","D1"]
chart.Categories.CategoryLabels = chart.ChartData["A2","A7"]
chart.Series[0].Values = chart.ChartData["B2","B7"]
chart.Series[0].Fill.FillType = FillFormatType.Solid
chart.Series[0].Fill.SolidColor.KnownColor = KnownColors.Brown
chart.Series[1].Values = chart.ChartData["C2","C7"]
chart.Series[1].Fill.FillType = FillFormatType.Solid
chart.Series[1].Fill.SolidColor.KnownColor = KnownColors.Green
chart.Series[2].Values = chart.ChartData["D2","D7"]
chart.Series[2].Fill.FillType = FillFormatType.Solid
chart.Series[2].Fill.SolidColor.KnownColor = KnownColors.Orange

#Set the 3D rotation
chart.RotationThreeD.XDegree = 10
chart.RotationThreeD.YDegree = 10

#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()