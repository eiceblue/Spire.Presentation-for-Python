from spire.presentation.common import *
from spire.presentation import *


inputFile = "Data/FormatChartDataLabels.pptx"
outputFile = "FormatChartDataLabels.pptx"

#Create PPT document and load file.
ppt = Presentation()
ppt.LoadFromFile(inputFile)

#Get the chart
chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Get the chart series
sers = chart.Series

#Initialize four instances of series label and set parameters of each label
cd1 = sers[0].DataLabels.Add()
cd1.PercentageVisible = True
cd1.TextFrame.Text = "Custom Datalabel1"
cd1.TextFrame.TextRange.FontHeight = 12
cd1.TextFrame.TextRange.LatinFont = TextFont("Lucida Sans Unicode")
cd1.TextFrame.TextRange.Fill.FillType =FillFormatType.Solid
cd1.TextFrame.TextRange.Fill.SolidColor.Color = Color.get_Green()

cd2 = sers[0].DataLabels.Add()
cd2.Position = ChartDataLabelPosition.InsideEnd
cd2.PercentageVisible = True
cd2.TextFrame.Text = "Custom Datalabel2"
cd2.TextFrame.TextRange.FontHeight = 10
cd2.TextFrame.TextRange.LatinFont = TextFont("Arial")
cd2.TextFrame.TextRange.Fill.FillType = FillFormatType.Solid
cd2.TextFrame.TextRange.Fill.SolidColor.Color = Color.get_OrangeRed()

cd3 = sers[0].DataLabels.Add()
cd3.Position = ChartDataLabelPosition.Center
cd3.PercentageVisible = True
cd3.TextFrame.Text = "Custom Datalabel3"
cd3.TextFrame.TextRange.FontHeight = 14
cd3.TextFrame.TextRange.LatinFont = TextFont("Calibri")
cd3.TextFrame.TextRange.Fill.FillType = FillFormatType.Solid
cd3.TextFrame.TextRange.Fill.SolidColor.Color = Color.get_Blue()

cd4 = sers[0].DataLabels.Add()
cd4.Position = ChartDataLabelPosition.InsideBase
cd4.PercentageVisible = True
cd4.TextFrame.Text = "Custom Datalabel4"
cd4.TextFrame.TextRange.FontHeight = 12
cd4.TextFrame.TextRange.LatinFont = TextFont("Lucida Sans Unicode")
cd4.TextFrame.TextRange.Fill.FillType = FillFormatType.Solid
cd4.TextFrame.TextRange.Fill.SolidColor.Color = Color.get_OliveDrab()

#Save and launch the file 
ppt.SaveToFile(outputFile, FileFormat.Pptx2010)
ppt.Dispose()
