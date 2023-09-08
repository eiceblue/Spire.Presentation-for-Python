from spire.presentation.common import *
from spire.presentation import *

inputFile = "./Data/Conversion.pptx"
outputFile = "ToXPS.xps"


#Create an instance of presentation document
ppt = Presentation()
#Load file
ppt.LoadFromFile(inputFile)

#Save the the XPS file
ppt.SaveToFile(outputFile, FileFormat.XPS)
ppt.Dispose()

      

    
