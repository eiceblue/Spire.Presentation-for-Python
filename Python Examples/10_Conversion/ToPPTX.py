from spire.presentation.common import *
from spire.presentation import *


inputFile = "./Data/ToPPTX.ppt"
outputFile = "ToPPTX.pptx"


#Create PPT document
presentation = Presentation()

#Load the PPT file from disk
presentation.LoadFromFile(inputFile)

#Save the PPT document to PPTX file format
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()

     


    
