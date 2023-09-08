from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/CopyParagraph.pptx"
outputFile = "ConvertPPTToOFD.ofd"

#Create an instance of presentation document
ppt = Presentation()
#Load file
ppt.LoadFromFile(inputFile)

ppt.SaveToFile(outputFile, FileFormat.OFD)
ppt.Dispose()

