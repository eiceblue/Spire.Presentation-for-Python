from spire.presentation import *

inputFile = "./Data/RemoveAllDigitalSignatures.pptx"
outputFile = "RemoveAllDigitalSignatures_out.pptx"

#Create a PowerPoint document.
ppt = Presentation()
#Load the file from disk.
ppt.LoadFromFile(inputFile)
#Remove all digital signatures
if ppt.IsDigitallySigned == True:
    ppt.RemoveAllDigitalSignatures()
#Save to file.
ppt.SaveToFile(outputFile, FileFormat.Pptx2010)
ppt.Dispose()