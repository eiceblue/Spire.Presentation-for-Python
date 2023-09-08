from spire.presentation.common import *
from spire.presentation import *


inputFile_1 = "Data/Template_Ppt_2.pptx"
inputFile_2 = "Data/Template_Ppt_1.pptx"
outputFile = "CopyChartBetweenPptFiles.pptx"

#Create a PPT document
presentation1 = Presentation()

#Load the file from disk.
presentation1.LoadFromFile(inputFile_1)

#Get the chart that is going to be copied.
chart = presentation1.Slides[0].Shapes[0] if isinstance(presentation1.Slides[0].Shapes[0], IChart) else None

#Load the second PowerPoint document.
presentation2 = Presentation()
presentation2.LoadFromFile(inputFile_2)

#Copy chart from the first document to the second document.
presentation2.Slides.Append()
presentation2.Slides[1].Shapes.CreateChart(chart, RectangleF.FromLTRB (100, 100, 600, 400), -1)

#Save to file.
presentation2.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation2.Dispose()
presentation1.Dispose()