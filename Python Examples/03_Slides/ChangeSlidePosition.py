from spire.presentation.common import *
from spire.presentation import *



inputFile ="./Data/ChangeSlidePosition.pptx"
outputFile ="ChangeSlidePosition.pptx"
#Create a PPT document
presentation = Presentation()
#Load the document from disk
presentation.LoadFromFile(inputFile)
#Move the first slide to the second slide position
slide = presentation.Slides[0]
slide.SlideNumber = 2
#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()
