from spire.presentation import *

inputFile = "./Data/Encrypt.pptx"
outputFile = "Encrypt_out.pptx"

#Create a PPT document
presentation = Presentation()
#Load the document from disk
presentation.LoadFromFile(inputFile)
#Get the password that the user entered
password = "e-iceblue"
#Encrypy the document with the password
presentation.Encrypt(password)
#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()