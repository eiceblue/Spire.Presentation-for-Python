from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/SomePresentation.pptx"
outputFile ="HighlightSpecifiedText.pptx"

ppt = Presentation()
ppt.LoadFromFile(inputFile)

#Get the specified shape
shape = ppt.Slides[0].Shapes[1]

options = TextHighLightingOptions()
options.WholeWordsOnly = True
options.CaseSensitive = True

shape.TextFrame.HighLightText("Spire", Color.get_Yellow(), options)

#Save to file.
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()
