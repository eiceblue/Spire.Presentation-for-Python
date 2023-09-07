from spire.presentation.common import *
from spire.presentation import *



inputFile ="./Data/ChangeSlidePosition.pptx"
outputFile ="SetStartingNumberForSlides.pptx"
#Create PPT document
presentation = Presentation()
#Load the PPT document from disk.
presentation.LoadFromFile(inputFile)
#Set 5 as the starting number
presentation.FirstSlideNumber = 5
#Save file
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()