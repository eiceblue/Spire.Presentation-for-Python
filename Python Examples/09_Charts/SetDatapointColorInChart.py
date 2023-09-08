from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/SetDatapointColorInChart.pptx"
outputFile = "SetDatapointColorInChart.pptx"

#Create PPT document and load file
ppt = Presentation()
ppt.LoadFromFile(inputFile)

#Get the chart
chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Initialize an instances of dataPoint
cdp1 = ChartDataPoint(chart.Series[0])

#Specify the datapoint order
cdp1.Index = 0

#Set the color of the datapoint
cdp1.Fill.FillType = FillFormatType.Solid
cdp1.Fill.SolidColor.KnownColor = KnownColors.Orange

#Add the dataPoint to first series
chart.Series[0].DataPoints.Add(cdp1)

#Set the color for the other three data points
cdp2 = ChartDataPoint(chart.Series[0])
cdp2.Index = 1
cdp2.Fill.FillType = FillFormatType.Solid
cdp2.Fill.SolidColor.KnownColor = KnownColors.Gold
chart.Series[0].DataPoints.Add(cdp2)

cdp3 = ChartDataPoint(chart.Series[0])
cdp3.Index = 2
cdp3.Fill.FillType = FillFormatType.Solid
cdp3.Fill.SolidColor.KnownColor = KnownColors.MediumPurple
chart.Series[0].DataPoints.Add(cdp3)

cdp4 = ChartDataPoint(chart.Series[0])
cdp4.Index = 1
cdp4.Fill.FillType = FillFormatType.Solid
cdp4.Fill.SolidColor.KnownColor = KnownColors.ForestGreen
chart.Series[0].DataPoints.Add(cdp4)


ppt.SaveToFile(outputFile, FileFormat.Pptx2010)
ppt.Dispose()

