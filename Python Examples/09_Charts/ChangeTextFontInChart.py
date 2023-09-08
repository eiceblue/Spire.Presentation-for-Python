from spire.presentation.common import *
from spire.presentation import *


inputFile = "Data/ChangeTextFontInChart.pptx"
outputFile = "ChangeTextFontInChart.pptx"

#Load a PPTX file
ppt = Presentation()
ppt.LoadFromFile(inputFile)

#Get the chart
chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Change the font of title
chart.ChartTitle.TextProperties.Paragraphs[0].DefaultCharacterProperties.LatinFont = TextFont("Lucida Sans Unicode")
chart.ChartTitle.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.SolidColor.KnownColor = KnownColors.Blue
chart.ChartTitle.TextProperties.Paragraphs[0].DefaultCharacterProperties.FontHeight = 30

#Change the font of legend
chart.ChartLegend.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.SolidColor.KnownColor = KnownColors.DarkGreen
chart.ChartLegend.TextProperties.Paragraphs[0].DefaultCharacterProperties.LatinFont = TextFont("Lucida Sans Unicode")

#Change the font of series
chart.PrimaryCategoryAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.SolidColor.KnownColor = KnownColors.Red
chart.PrimaryCategoryAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.FillType = FillFormatType.Solid
chart.PrimaryCategoryAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.FontHeight = 10
chart.PrimaryCategoryAxis.TextProperties.Paragraphs[0].DefaultCharacterProperties.LatinFont = TextFont("Lucida Sans Unicode")
ppt.SaveToFile(outputFile, FileFormat.Pptx2010)
ppt.Dispose()


