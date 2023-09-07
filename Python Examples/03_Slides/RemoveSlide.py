from spire.presentation.common import *
from spire.presentation import *


inputFile ="./Data/RemoveSlide.pptx"
outputFile ="RemoveSlide.pptx"
presentation = Presentation()
presentation.LoadFromFile(inputFile)
#Remove slide by index
presentation.Slides.RemoveAt(0)
#Remove slide by its reference
slide = presentation.Slides[1]
presentation.Slides.Remove(slide)
#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()
