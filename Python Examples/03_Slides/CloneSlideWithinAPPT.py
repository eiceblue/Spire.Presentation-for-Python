from spire.presentation.common import *
from spire.presentation import *


inputFile ="./Data/InputTemplate.pptx"
outputFile ="CloneSlideWithinAPPT.pptx"

#Create an instance of presentation document
ppt = Presentation()
#Load file
ppt.LoadFromFile(inputFile)
#Get a list of slides and choose the first slide to be cloned
slide = ppt.Slides[0]
#Insert the desired slide to the specified index in the same presentation
index = 1
ppt.Slides.Insert(index, slide)
#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()
