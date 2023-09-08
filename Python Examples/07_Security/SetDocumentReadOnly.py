from spire.presentation import *

inputFile = "./Data/InputTemplate.pptx"
outputFile = "SetDocumentReadOnly_out.pptx"

#Load a PPT document
presentation = Presentation()
#Load the document from disk
presentation.LoadFromFile(inputFile)
#Get the password that the user entered
password = "e-iceblue"
#Protect the document with the password
presentation.Protect(password)
#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()