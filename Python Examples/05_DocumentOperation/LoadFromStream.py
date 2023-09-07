from spire.presentation import *

inputFile = "./Data/InputTemplate.pptx"
outputFile = "LoadFromStream.pptx"

#Create an instance of presentation document
ppt = Presentation()
#Load PowerPoint file from stream
from_stream = Stream(inputFile)
ppt.LoadFromStream(from_stream, FileFormat.Pptx2013)
#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
from_stream.Dispose()
ppt.Dispose()
