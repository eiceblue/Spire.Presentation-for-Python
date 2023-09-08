from spire.presentation.common import *
from spire.presentation import *


inputFile = "Data/Template_Ppt_2.pptx"
outputFile = "HideOrShowASeriesOfChart.pptx"

#Create a PowerPoint document.
presentation = Presentation()

#Load the file from disk.
presentation.LoadFromFile(inputFile)

#Get the first slide.
slide = presentation.Slides[0]

#Get the first chart.
chart = slide.Shapes[0] if isinstance(slide.Shapes[0], IChart) else None

#Hide the first series of the chart.
chart.Series[0].IsHidden = True

#Show the first series of the chart.
#chart.Series[0].IsHidden = false

#Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()
