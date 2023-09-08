from spire.presentation.common import *
from spire.presentation import *


inputFile = "./Data/AddAndFormatErrorBars.pptx"
outputFile = "AddAndFormatErrorBars.pptx"

#Create a PowerPoint document.
presentation = Presentation()

#Load the file from disk.
presentation.LoadFromFile(inputFile)

#Get the column chart on the first slide and set chart title.
columnChart = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IChart) else None
columnChart.ChartTitle.TextProperties.Text = "Vertical Error Bars"

#Add Y (Vertical) Error Bars.
#Get Y error bars of the first chart series.
errorBarsYFormat1 = columnChart.Series[0].ErrorBarsYFormat

#Set end cap.
errorBarsYFormat1.ErrorBarNoEndCap = False

#Specify direction.
errorBarsYFormat1.ErrorBarSimType = ErrorBarSimpleType.Plus

#Specify error amount type.
errorBarsYFormat1.ErrorBarvType = ErrorValueType.StandardError

#Set value.
errorBarsYFormat1.ErrorBarVal = 0.3

#Set line format.
errorBarsYFormat1.Line.FillType = FillFormatType.Solid
errorBarsYFormat1.Line.SolidFillColor.Color = Color.get_MediumVioletRed()
errorBarsYFormat1.Line.Width = 1

#Get the bubble chart on the second slide and set chart title.
bubbleChart = presentation.Slides[1].Shapes[0] if isinstance(presentation.Slides[1].Shapes[0], IChart) else None
bubbleChart.ChartTitle.TextProperties.Text = "Vertical and Horizontal Error Bars"

#Add X (Horizontal) and Y (Vertical) Error Bars.
#Get X error bars of the first chart series.
errorBarsXFormat = bubbleChart.Series[0].ErrorBarsXFormat

#Set end cap.
errorBarsXFormat.ErrorBarNoEndCap = False

#Specify direction.
errorBarsXFormat.ErrorBarSimType = ErrorBarSimpleType.Both

#Specify error amount type.
errorBarsXFormat.ErrorBarvType = ErrorValueType.StandardError

#Set value.
errorBarsXFormat.ErrorBarVal = 0.3

#Get Y error bars of the first chart series.
errorBarsYFormat2 = bubbleChart.Series[0].ErrorBarsYFormat

#Set end cap.
errorBarsYFormat2.ErrorBarNoEndCap = False

#Specify direction.
errorBarsYFormat2.ErrorBarSimType = ErrorBarSimpleType.Both

#Specify error amount type.
errorBarsYFormat2.ErrorBarvType = ErrorValueType.StandardError

#Set value.
errorBarsYFormat2.ErrorBarVal = 0.3

#Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()
