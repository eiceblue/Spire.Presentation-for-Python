from spire.presentation import *

inputFile = "./Data/Properties.pptx"
outputFile = "Properties_out.pptx"
     
#Create a PPT document
presentation = Presentation()
presentation.LoadFromFile(inputFile)
#Set the DocumentProperty of PPT document
presentation.DocumentProperty.Application = "Spire.Presentation"
presentation.DocumentProperty.Author = "E-iceblue"
presentation.DocumentProperty.Company = "E-iceblue Co., Ltd."
presentation.DocumentProperty.Keywords = "Demo File"
presentation.DocumentProperty.Comments = "This file is used to test Spire.Presentation."
presentation.DocumentProperty.Category = "Demo"
presentation.DocumentProperty.Title = "This is a demo file."
presentation.DocumentProperty.Subject = "Test"
#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()