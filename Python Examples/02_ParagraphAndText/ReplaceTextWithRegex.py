from spire.presentation.common import *
from spire.presentation import *


inputFile ="./Data/SomePresentation.pptx"
outputFile ="ReplaceTextWithRegex.pptx"

#Create Presentation
presentation = Presentation()

#Load file
presentation.LoadFromFile(inputFile)

#Regex for all words
regex = Regex("\\d+.\\d+|\\w+")

#New string value
newvalue = "This is the test!"

#Loop and replace
for slide in presentation.Slides:
    for shape in slide.Shapes:
        shape.ReplaceTextWithRegex(regex, newvalue)

#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()