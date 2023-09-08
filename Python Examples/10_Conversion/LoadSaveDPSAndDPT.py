from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/sample.dps"
outputFile = "LoadSaveDPSAndDPT.dps"
outputFile2 = "LoadSaveDPSAndDPT.dpt"
#Create PPT document
presentation = Presentation()

#Load the PPT document from disk.
presentation.LoadFromFile(inputFile, FileFormat.Dps)

presentation.SaveToFile(outputFile, FileFormat.Dps)
presentation.SaveToFile(outputFile2, FileFormat.Dpt)
presentation.Dispose()

