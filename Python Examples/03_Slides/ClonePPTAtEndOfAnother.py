from spire.presentation.common import *
from spire.presentation import *


inputFile_1 = "./Data/ChangeSlidePosition.pptx"
inputFile_2 = "./Data/PPTSample_N.pptx"
outputFile ="ClonePPTAtEndOfAnother.pptx"
#Load source document from disk
sourcePPT = Presentation()
sourcePPT.LoadFromFile(inputFile_1)
#Load destination document from disk
destPPT = Presentation()
destPPT.LoadFromFile(inputFile_2)
#Loop through all slides of source document
for slide in sourcePPT.Slides:
    #Append the slide at the end of destination document
    destPPT.Slides.AppendBySlide(slide)
#Save the document
destPPT.SaveToFile(outputFile, FileFormat.Pptx2013)
destPPT.Dispose()
