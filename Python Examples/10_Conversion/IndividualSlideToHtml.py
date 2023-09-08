from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/ChangeSlidePosition.pptx"
outputFile = "IndividualSlideToHtml.pptx"

#Create PPT document
presentation = Presentation()

#Load the PPT document from disk.
presentation.LoadFromFile(inputFile)

#Get the first slide
slide = presentation.Slides[0]

#String for output file 

#Save the first slide to HTML 
slide.SaveToFile(outputFile, FileFormat.Html)
presentation.Dispose()

