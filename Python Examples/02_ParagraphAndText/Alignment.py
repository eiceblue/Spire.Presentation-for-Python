from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/Alignment.pptx"
outputFile = "Alignment.pptx"

#Create a PPT document and load file
presentation = Presentation()
presentation.LoadFromFile(inputFile)

#Get the related shape and set the text alignment
shape = presentation.Slides[0].Shapes[1]
shape.TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Left
shape.TextFrame.Paragraphs[1].Alignment = TextAlignmentType.Center
shape.TextFrame.Paragraphs[2].Alignment = TextAlignmentType.Right
shape.TextFrame.Paragraphs[3].Alignment = TextAlignmentType.Justify
shape.TextFrame.Paragraphs[4].Alignment = TextAlignmentType.none

#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()

