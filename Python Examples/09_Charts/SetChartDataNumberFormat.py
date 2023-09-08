from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/SetChartDataNumberFormat.pptx"
outputFile = "SetChartDataNumberFormat.pptx"

#Create PPT document and load file
presentation = Presentation()
presentation.LoadFromFile(inputFile)

#Get chart on the first slide
chart = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IChart) else None

#Set the number format for Axis
chart.PrimaryValueAxis.NumberFormat = "#,##0.00"

#Set the DataLabels format for Axis
chart.Series[0].DataLabels.LabelValueVisible = True
chart.Series[0].DataLabels.PercentValueVisible = False
chart.Series[0].DataLabels.NumberFormat = "#,##0.00"
chart.Series[0].DataLabels.HasDataSource = False

#Set the number format for ChartData
for i in range(1, (chart.Series[0].Values.Count) + 1):
    chart.ChartData[i,1].NumberFormat = "#,##0.00"

#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()

