from spire.presentation.common import *
from spire.presentation import *


inputFile ="./Data/SetColorAndNameForTrendline.pptx"
outputFile = "SetColorAndNameForTrendline.pptx"

#Create an instance of presentation document
ppt = Presentation()
#Load file
ppt.LoadFromFile(inputFile)
#Find the first chart in the first Slide
chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Find the first trendline in the chart
trendline = chart.Series[0].TrendLines[0] if isinstance(chart.Series[0].TrendLines[0], ITrendlines) else None

#Set name for trendline
trendline.Name = "trendlineName"

#Set color for trendline
trendline.Line.FillType = FillFormatType.Solid
trendline.Line.SolidFillColor.Color = Color.get_Red()

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2010)
ppt.Dispose()


