from spire.presentation import *

inputFile = "./Data/Template_Ppt_4.pptx"
outputFile = "ModifyPasswordOfEncryptedPPT.pptx"

#Create a PowerPoint document.
presentation = Presentation()
#Load the file from disk.
presentation.LoadFromFile(inputFile, "123456")
#Remove the encryption.
presentation.RemoveEncryption()
#Protect the document by setting a new password.
presentation.Protect("654321")
#Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()