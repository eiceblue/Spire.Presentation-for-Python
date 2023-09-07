from spire.presentation.common import *
from spire.presentation import *


inputFile_1 = "./Data/CloneSlideToAnotherPPT-1.pptx"
inputFile_2 = "./Data/CloneSlideToAnotherPPT-2.pptx"
outputFile ="CloneSlideToAnotherPPT.pptx"
#Create a PPT document
presentation = Presentation()
#Load the document from disk
presentation.LoadFromFile(inputFile_2)
#Load the another document and choose the first slide to be cloned
ppt1 = Presentation()
ppt1.LoadFromFile(inputFile_1)
slide1 = ppt1.Slides[0]
#Insert the slide to the specified index in the source presentation
index = 1
presentation.Slides.Insert(index, slide1)
#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()