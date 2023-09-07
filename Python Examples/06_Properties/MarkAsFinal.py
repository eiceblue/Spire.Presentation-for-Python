from spire.presentation import *

inputFile = "./Data/MarkAsFinal.pptx"
outputFile = "MarkAsFinal_out.pptx"

#Create a PPT document
presentation = Presentation()
#Load the document from disk
presentation.LoadFromFile(inputFile)
#Mark the document as final
presentation.DocumentProperty.MarkAsFinal = True
#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()