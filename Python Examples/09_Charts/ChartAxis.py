from spire.presentation.common import *
from spire.presentation import *


inputFile ="Data/ChartAxis.pptx"
outputFile = "ChartAxis.pptx"

#Create PPT document and load file
ppt = Presentation()
ppt.LoadFromFile(inputFile)

#Get the chart
chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Add a secondary axis to display the value of Series 3
chart.Series[2].UseSecondAxis = True

#Set the grid line of secondary axis as invisible
chart.SecondaryValueAxis.MajorGridTextLines.FillType = FillFormatType.none

#Set bounds of axis value. Before we assign values, we must set IsAutoMax and IsAutoMin as false, otherwise MS PowerPoint will automatically set the values.
chart.PrimaryValueAxis.IsAutoMax = False
chart.PrimaryValueAxis.IsAutoMin = False
chart.SecondaryValueAxis.IsAutoMax = False
chart.SecondaryValueAxis.IsAutoMax = False
chart.PrimaryValueAxis.MinValue = 0
chart.PrimaryValueAxis.MaxValue = 5.0
chart.SecondaryValueAxis.MinValue = 0
chart.SecondaryValueAxis.MaxValue = 1.0

#Set axis line format
chart.PrimaryValueAxis.MinorGridLines.FillType = FillFormatType.Solid
chart.SecondaryValueAxis.MinorGridLines.FillType = FillFormatType.Solid
chart.PrimaryValueAxis.MinorGridLines.Width = 0.1
chart.SecondaryValueAxis.MinorGridLines.Width = 0.1
chart.PrimaryValueAxis.MinorGridLines.SolidFillColor.Color = Color.get_LightGray()
chart.SecondaryValueAxis.MinorGridLines.SolidFillColor.Color = Color.get_LightGray()
chart.PrimaryValueAxis.MinorGridLines.DashStyle = LineDashStyleType.Dash
chart.SecondaryValueAxis.MinorGridLines.DashStyle = LineDashStyleType.Dash
chart.PrimaryValueAxis.MajorGridTextLines.Width = 0.3
chart.PrimaryValueAxis.MajorGridTextLines.SolidFillColor.Color = Color.get_LightSkyBlue()
chart.SecondaryValueAxis.MajorGridTextLines.Width = 0.3
chart.SecondaryValueAxis.MajorGridTextLines.SolidFillColor.Color = Color.get_LightSkyBlue()

ppt.SaveToFile(outputFile, FileFormat.Pptx2010)
ppt.Dispose()