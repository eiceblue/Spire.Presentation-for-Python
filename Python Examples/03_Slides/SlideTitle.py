from spire.presentation.common import *
from spire.presentation import *


inputFile ="./Data/InputTemplate.pptx"
outputFile ="SlideTitle.pptx"
ppt = Presentation()
ppt.LoadFromFile(inputFile)
#Get the first slide
slide = ppt.Slides[0]
#Get the title of the first slide
slideTitle = slide.Title
#Set the title of the second slide
ppt.Slides[1].Title = "Second Slide"
#Save to file.
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()