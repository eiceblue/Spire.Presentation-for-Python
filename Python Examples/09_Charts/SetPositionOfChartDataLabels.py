from spire.presentation.common import *
from spire.presentation import *


inputFile ="./Data/Template_Ppt_2.pptx"
outputFile = "SetPositionOfChartDataLabels.pptx"

#Create a PowerPoint document.
presentation = Presentation()

#Load the file from disk.
presentation.LoadFromFile(inputFile)

#Get the chart.
chart = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IChart) else None

#Add data label to chart and set its id.
label1 = chart.Series[0].DataLabels.Add()
label1.ID = 0

#Set the default position of data label. This position is relative to the data markers.
#label1.Position = ChartDataLabelPosition.OutsideEnd

#Set custom position of data label. This position is relative to the default position.
label1.X = 0.1
label1.Y = -0.1

#Set label value visible
label1.LabelValueVisible = True

#Set legend key invisible
label1.LegendKeyVisible = False

#Set category name invisible
label1.CategoryNameVisible = False

#Set series name invisible
label1.SeriesNameVisible = False

#Set Percentage invisible
label1.PercentageVisible = False

#Set border style and fill style of data label
label1.Line.FillType = FillFormatType.Solid
label1.Line.SolidFillColor.Color = Color.get_Blue()
label1.Fill.FillType = FillFormatType.Solid
label1.Fill.SolidColor.Color = Color.get_Orange()

#Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()
