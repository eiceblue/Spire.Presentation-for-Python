from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/Conversion.pptx"
outputFile = "ToHTML.html"

#Create an instance of presentation document
ppt = Presentation()

#Load file
ppt.LoadFromFile(inputFile)

#Save the document to HTML format
ppt.SaveToFile(outputFile, FileFormat.Html)
ppt.Dispose()

