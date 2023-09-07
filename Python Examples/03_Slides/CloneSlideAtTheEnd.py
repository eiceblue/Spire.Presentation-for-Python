from spire.presentation.common import *
from spire.presentation import *


inputFile ="./Data/ChangeSlidePosition.pptx"
outputFile ="CloneSlideAtTheEnd.pptx"
#Load PPT document from disk
presentation = Presentation()
presentation.LoadFromFile(inputFile)
#Get the first slide
slide = presentation.Slides[0]
#Append the slide at the end of the document
presentation.Slides.AppendBySlide(slide)
#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()