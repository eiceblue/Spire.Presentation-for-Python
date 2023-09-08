from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/Template_Ppt_3.pptx"
outputFile = "RemoveChartFromPptSlide.pptx"

#Create a PowerPonit document
presentation = Presentation()

#Load the file from disk.
presentation.LoadFromFile(inputFile)

#Get the first slide from the document.
slide = presentation.Slides[0]

#Remove chart from the slide.
for i, unusedItem in enumerate(slide.Shapes):
    shape = slide.Shapes[i] if isinstance(slide.Shapes[i], IShape) else None
    if isinstance(shape, IChart):
        slide.Shapes.Remove(shape)

#Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()

