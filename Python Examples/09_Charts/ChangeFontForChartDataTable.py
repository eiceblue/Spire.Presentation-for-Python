from spire.presentation.common import *
from spire.presentation import *


inputFile = "Data/ChartSample2.pptx"
outputFile = "ChangeFontSizeForChartDataTable.pptx"

#Create PPT document and load file
ppt = Presentation()
ppt.LoadFromFile(inputFile)

#Get chart on the first slide
Chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None
Chart.HasDataTable = True

#Add a new paragraph in data table
tp = TextParagraph()
Chart.ChartDataTable.Text.Paragraphs.Append(tp)

#Change the font size
Chart.ChartDataTable.Text.Paragraphs[0].DefaultCharacterProperties.FontHeight = 15

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2010)
ppt.Dispose()
