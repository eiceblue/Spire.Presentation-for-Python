from spire.presentation.common import *
from spire.presentation import *


inputFile = "./Data/ChartSample1.pptx"
outputFile = "AddCustomErrorBars.pptx"

#Create PPT document and load file
ppt = Presentation()
ppt.LoadFromFile(inputFile)

#Get the bubble chart on the first slide
bubbleChart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Get X error bars of the first chart series
errorBarsXFormat = bubbleChart.Series[0].ErrorBarsXFormat

#Specify error amount type as custom error bars
errorBarsXFormat.ErrorBarvType = ErrorValueType.CustomErrorBars

#Set the minus and plus value of the X error bars
errorBarsXFormat.MinusVal = 0.5
errorBarsXFormat.PlusVal = 0.5

#Get Y error bars of the first chart series
errorBarsYFormat = bubbleChart.Series[0].ErrorBarsYFormat

#Specify error amount type as custom error bars
errorBarsYFormat.ErrorBarvType = ErrorValueType.CustomErrorBars

#Set the minus and plus value of the Y error bars
errorBarsYFormat.MinusVal = 1
errorBarsYFormat.PlusVal = 1

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2010)
ppt.Dispose()