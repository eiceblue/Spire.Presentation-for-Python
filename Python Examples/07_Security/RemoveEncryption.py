from spire.presentation import *

inputFile = "./Data/Template_Ppt_4.pptx"
outputFile = "RemoveEncryption.pptx"


#Create a PowerPoint document.
presentation = Presentation()
#Load the file from disk.
presentation.LoadFromFile(inputFile, "123456")
#Remove encryption.
presentation.RemoveEncryption()
#Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()