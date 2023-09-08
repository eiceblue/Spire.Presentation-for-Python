from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/ChangeSlidePosition.pptx"
outputFile = "SpecificSlideToPDF.pdf"

#Create PPT document
presentation = Presentation()

#Load the PPT document from disk.
presentation.LoadFromFile(inputFile)

#Get the second slide
slide = presentation.Slides[1]

#String for output file 

#Save the second slide to PDF
slide.SaveToFile(outputFile,FileFormat.PDF)
presentation.Dispose()


