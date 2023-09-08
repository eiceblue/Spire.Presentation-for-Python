from spire.presentation.common import *
from spire.presentation import *


inputFile = "Data/TrendlineEquation.pptx"
outputFile = "ChangesForTrendLineEquation.pptx"

#Create Presentation
presentation = Presentation()

#Load ppt file
presentation.LoadFromFile(inputFile)

#Get chart on the first slide
chart = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IChart) else None

#Get the first trendline 
trendline = chart.Series[0].TrendLines[0]

#Change font size for trendline Equation text
for para in trendline.TrendLineLabel.TextFrameProperties.Paragraphs:
    para.DefaultCharacterProperties.FontHeight = 20
    for range in para.TextRanges:
        range.FontHeight = 20

#Change position for trendline Equation
trendline.TrendLineLabel.OffsetX = -0.1
trendline.TrendLineLabel.OffsetY = -0.05

#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()
