from spire.presentation.common import *
from spire.presentation import *


inputFile = "Data/Template_Ppt_2.pptx"
outputFile = "CopyChartWithinOnePPT.pptx"

ppt = Presentation()

#Load the file from disk.
ppt.LoadFromFile(inputFile)

#Get the chart that is going to be copied.
chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Copy the chart from the first slide to the specified location of the second slide within the same document.
slide1 = ppt.Slides.Append()
slide1.Shapes.CreateChart(chart, RectangleF.FromLTRB (100, 100, 600, 400), 0)

#Save to file.
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()