from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/VaryColorsOfSameSeriesDataMarkers.pptx"
outputFile = "VaryColorOfSameSerieDataMarkers.pptx"

#Create a PowerPoint document.
presentation = Presentation()

#Load the file from disk.
presentation.LoadFromFile(inputFile)

#Get the chart from the presentation.
chart = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IChart) else None

#Create a ChartDataPoint object and specify the index.
dataPoint = ChartDataPoint(chart.Series[0])
dataPoint.Index = 0

#Set the fill color of the data marker.
dataPoint.MarkerFill.Fill.FillType = FillFormatType.Solid
dataPoint.MarkerFill.Fill.SolidColor.Color = Color.get_Red()

#Set the line color of the data marker.
dataPoint.MarkerFill.Line.FillType = FillFormatType.Solid
dataPoint.MarkerFill.Line.SolidFillColor.Color = Color.get_Red()

#Add the data point to the point collection of a series.
chart.Series[0].DataPoints.Add(dataPoint)

dataPoint = ChartDataPoint(chart.Series[0])
dataPoint.Index = 1
dataPoint.MarkerFill.Fill.FillType = FillFormatType.Solid
dataPoint.MarkerFill.Fill.SolidColor.Color = Color.get_Black()
dataPoint.MarkerFill.Line.FillType = FillFormatType.Solid
dataPoint.MarkerFill.Line.SolidFillColor.Color = Color.get_Black()
chart.Series[0].DataPoints.Add(dataPoint)

dataPoint = ChartDataPoint(chart.Series[0])
dataPoint.Index = 2
dataPoint.MarkerFill.Fill.FillType = FillFormatType.Solid
dataPoint.MarkerFill.Fill.SolidColor.Color = Color.get_Blue()
dataPoint.MarkerFill.Line.FillType = FillFormatType.Solid
dataPoint.MarkerFill.Line.SolidFillColor.Color = Color.get_Blue()
chart.Series[0].DataPoints.Add(dataPoint)

#Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()

