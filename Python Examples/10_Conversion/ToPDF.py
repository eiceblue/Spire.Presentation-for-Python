from spire.presentation.common import *
from spire.presentation import *


inputFile ="./Data/ToPDF.pptx"
outputFile = "ToPDF.pdf"

#Create a PPT document
presentation = Presentation()

#Load PPT file from disk
presentation.LoadFromFile(inputFile)

#Save the PPT to PDF file format
presentation.SaveToFile(outputFile, FileFormat.PDF)
presentation.Dispose()

