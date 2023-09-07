from spire.presentation import *

outputFile_pptx = "SetPropertiesForTemplate.pptx"
outputFile_ppt = "SetPropertiesForTemplate.ppt"
outputFile_odp = "SetPropertiesForTemplate.odp"

def SetPropertiesForTemplate(filePath, fileFormat):
     #Create a document
     presentation = Presentation()
     #Set the DocumentProperty 
     presentation.DocumentProperty.Application = "Spire.Presentation"
     presentation.DocumentProperty.Author = "E-iceblue"
     presentation.DocumentProperty.Company = "E-iceblue Co., Ltd."
     presentation.DocumentProperty.Keywords = "Demo File"
     presentation.DocumentProperty.Comments = "This file is used to test Spire.Presentation."
     presentation.DocumentProperty.Category = "Demo"
     presentation.DocumentProperty.Title = "This is a demo file."
     presentation.DocumentProperty.Subject = "Test"
     #Save to template file
     presentation.SaveToFile(filePath, fileFormat)
     presentation.Dispose()


#Create the .pptx template
SetPropertiesForTemplate(outputFile_pptx, FileFormat.Pptx2013)
#Create the .odp template
SetPropertiesForTemplate(outputFile_odp, FileFormat.ODP)
#Create the .ppt template
SetPropertiesForTemplate(outputFile_ppt, FileFormat.PPT)