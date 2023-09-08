from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/Template_Ppt_3.pptx"
outputFile = "SetTextFontForChartTitle.pptx"

#Create a PowerPoint document.
presentation = Presentation()

#Load the file from disk.
presentation.LoadFromFile(inputFile)

#Get the chart.
chart = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IChart) else None

#Set the font for the text on chart title area.
chart.ChartTitle.TextProperties.Paragraphs[0].DefaultCharacterProperties.LatinFont = TextFont("Arial Unicode MS")
chart.ChartTitle.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.SolidColor.KnownColor = KnownColors.Blue
chart.ChartTitle.TextProperties.Paragraphs[0].DefaultCharacterProperties.FontHeight = 50

#Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()

