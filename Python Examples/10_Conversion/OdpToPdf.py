from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/toPdf.odp"
outputFile = "OdpToPdf.pdf"
#Create an instance of presentation document
ppt = Presentation()
#Load file
ppt.LoadFromFile(inputFile)
#Save the Odp document to PDF file format
ppt.SaveToFile(outputFile, FileFormat.PDF)
ppt.Dispose()
