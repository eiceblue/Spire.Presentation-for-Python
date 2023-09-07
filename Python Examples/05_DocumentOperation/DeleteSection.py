from spire.presentation import *

inputFile = "./Data/AddSection.pptx"
outputFile = "DeleteSection.pptx"

#Create a PPT document
ppt = Presentation()
ppt.LoadFromFile(inputFile)
ppt.SectionList.RemoveAll()
#Save to file.
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()