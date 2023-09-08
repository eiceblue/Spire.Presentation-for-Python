from spire.presentation.common import *
from spire.presentation import *


inputFile ="./Data/ChartSample2.pptx"
outputFile = "SetChartBorderstyle.pptx"

#Create Presentation
presentation = Presentation()

#Load ppt file
presentation.LoadFromFile(inputFile)

#Get chart on the first slide
chart = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IChart) else None

#Set border style
chart.Line.FillFormat.FillType = FillFormatType.Solid
chart.Line.FillFormat.SolidFillColor.Color = Color.get_Red()
chart.BorderRoundedCorners = True

#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()

