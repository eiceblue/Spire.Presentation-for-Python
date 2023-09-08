from spire.presentation.common import *
from spire.presentation import *


outputFile ="CreateFunnelChart.pptx"

#Create PPT document
ppt = Presentation()

#Create a Funnel chart to the first slide
chart = ppt.Slides[0].Shapes.AppendChartInit (ChartType.Funnel, RectangleF.FromLTRB (50, 50, 600, 450), False)

#Set series text
chart.ChartData[0,1].Text = "Series 1"

#Set category text
categories = ["Website Visits", "Download", "Uploads", "Requested price", "Invoice sent", "Finalized"]
i = 0
while i < len(categories):
    chart.ChartData[i + 1,0].Text = categories[i]
    i += 1

#Fill data for chart
values = [50000, 47000, 30000, 15000, 9000, 5600]
i = 0
while i < len(values):
    chart.ChartData[i + 1,1].NumberValue = values[i]
    i += 1

#Set series labels
chart.Series.SeriesLabel = chart.ChartData[0,1,0,1]

#Set categories labels 
chart.Categories.CategoryLabels = chart.ChartData[1,0,len(categories),0]

#Assign data to series values
chart.Series[0].Values = chart.ChartData[1,1,len(values),1]

#Set the chart title
chart.ChartTitle.TextProperties.Text = "Funnel"

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()
