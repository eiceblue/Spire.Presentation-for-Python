
from spire.presentation.common import *
from spire.presentation import *

inputFile_1 = "./Data/TextTemplate.pptx"
inputFile_2 = "./Data/CopyParagraph.pptx"
outputFile ="CopyParagraphToAnotherPPT.pptx"
#Load the source file
ppt1 = Presentation()
ppt1.LoadFromFile(inputFile_1)

#Get the text from the first shape on the first slide
sourceshp = ppt1.Slides[0].Shapes[0]
text = (sourceshp).TextFrame.Text

#Load the target file
ppt2 = Presentation()
ppt2.LoadFromFile(inputFile_2)

#Get the first shape on the first slide from the target file
destshp = ppt2.Slides[0].Shapes[0]

#Add the text to the target file
(destshp).TextFrame.Text += "\n\n" + text

#Save the document
ppt2.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt2.Dispose()