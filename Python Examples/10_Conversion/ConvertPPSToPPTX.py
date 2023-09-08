from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/Conversion.pps"
outputFile = "ConvertPPSToPPTX.pptx"

#Create an instance of presentation document
ppt = Presentation()
#Load file
ppt.LoadFromFile(inputFile)

#Save the PPS document to PPTX file format
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()


