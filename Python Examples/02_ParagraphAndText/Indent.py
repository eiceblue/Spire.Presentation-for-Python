from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/Indent.pptx"
outputFile ="Indent.pptx"

#Load a PPT document
presentation = Presentation()
presentation.LoadFromFile(inputFile)

shape = presentation.Slides[0].Shapes[0]
paras = shape.TextFrame.Paragraphs

#Set the paragraph style for first paragraph
paras[0].Indent = 20
paras[0].LeftMargin = 10
paras[0].SpaceAfter = 10

#Set the paragraph style of the third paragraph 
paras[2].Indent = -100
paras[2].LeftMargin = 40
paras[2].SpaceBefore = 0
paras[2].SpaceAfter = 0

#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()