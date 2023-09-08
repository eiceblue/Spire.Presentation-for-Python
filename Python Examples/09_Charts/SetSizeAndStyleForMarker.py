from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/SetSizeAndStyleForMarker.pptx"
outputFile = "SetSizeAndStyleForMarker.pptx"

#Create an instance of presentation document
ppt = Presentation()
#Load file
ppt.LoadFromFile(inputFile)
chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

for i, unusedItem in enumerate(chart.Series[0].Values):
    #Create a ChartDataPoint object and specify the index.
    dataPoint = ChartDataPoint(chart.Series[0])
    dataPoint.Index = i

    #Set the fill color of the data marker.
    dataPoint.MarkerFill.Fill.FillType = FillFormatType.Solid
    dataPoint.MarkerFill.Fill.SolidColor.Color = Color.get_Yellow()

    #Set the line color of the data marker.
    dataPoint.MarkerFill.Line.FillType = FillFormatType.Solid
    dataPoint.MarkerFill.Line.SolidFillColor.Color = Color.get_YellowGreen()

    #Set the size of the data marker.
    dataPoint.MarkerSize = 20

    #Set the style of the data marker
    dataPoint.MarkerStyle = ChartMarkerType.Diamond
    chart.Series[0].DataPoints.Add(dataPoint)
#Save to file.
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()

