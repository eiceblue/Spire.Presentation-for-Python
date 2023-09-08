from spire.presentation.common import *
from spire.presentation import *

def _GetTotal(ranges):
    total = 0
    for i, unusedItem in enumerate(ranges):
        total += float(ranges[i].Text)
    return total


inputFile ="./Data/ColumnStacked.pptx"
outputFile = "SetPercentageForLabels.pptx"

#Create PPT document and load file
ppt = Presentation()
ppt.LoadFromFile(inputFile)

#Get the chart on the first slide
Chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

dataPontPercent = 0

for i, unusedItem in enumerate(Chart.Series):
    series = Chart.Series[i]
    #Get the total number
    total = _GetTotal(series.Values)
    for j, unusedItem in enumerate(series.Values):
        #Get the percent
        dataPontPercent = float(series.Values[j].Text) / total * 100
        #Add datalabels
        label = series.DataLabels.Add()
        label.LabelValueVisible = True
        #Set the percent text for the label
        label.TextFrame.Paragraphs[0].Text = "{0:.2F} %".format(dataPontPercent)
        label.TextFrame.Paragraphs[0].TextRanges[0].FontHeight = 12

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2010)
ppt.Dispose()



