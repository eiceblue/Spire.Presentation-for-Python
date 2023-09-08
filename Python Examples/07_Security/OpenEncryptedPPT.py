from spire.presentation import *

inputFile = "./Data/OpenEncryptedPPT.pptx"
outputFile = "OpenEncryptedPPT_out.pptx"

#Create a PPT document
presentation = Presentation()
#Load the PPT with password
presentation.LoadFromFile(inputFile, FileFormat.Pptx2010, "123456")
#Save as a new PPT with original password
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()