from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/Template_Ppt_2.pptx"
outputFile = "SetTextFontForLegendAndAxis.pptx"

#Create a PowerPonit document.
presentation = Presentation()

#Load the file from disk.
presentation.LoadFromFile(inputFile)

#Get the chart.
chart = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IChart) else None

#Set the font for the text on Chart Legend area.
chart.ChartLegend.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.SolidColor.KnownColor = KnownColors.Green
chart.ChartLegend.TextProperties.Paragraphs[0].DefaultCharacterProperties.LatinFont = TextFont("Arial Unicode MS")

#Set the font for the text on Chart Axis area.
chart.PrimaryCategoryAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.SolidColor.KnownColor = KnownColors.Red
chart.PrimaryCategoryAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.FillType = FillFormatType.Solid
chart.PrimaryCategoryAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.FontHeight = 10
chart.PrimaryCategoryAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.LatinFont = TextFont("Arial Unicode MS")

#Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()

