from spire.presentation.common import *
from spire.presentation import *


inputFile = "Data/ChartSample2.pptx"
outputFile = "ChangeColorForLegend.pptx"

#Create PPT document and load file
ppt = Presentation()
ppt.LoadFromFile(inputFile)

#Get chart on the first slide
Chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Change the fill color
Chart.ChartLegend.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.FillType = FillFormatType.Solid
Chart.ChartLegend.TextProperties.Paragraphs[0].DefaultCharacterProperties.Fill.SolidColor.Color = Color.get_Blue()

#Use italic for the paragraph
Chart.ChartLegend.TextProperties.Paragraphs[0].DefaultCharacterProperties.IsItalic = TriState.TTrue

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2010)
ppt.Dispose()

