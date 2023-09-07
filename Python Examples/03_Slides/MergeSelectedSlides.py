from spire.presentation.common import *
from spire.presentation import *


inputFile_1 = "./Data/InputTemplate.pptx"
inputFile_2 = "./Data/TextTemplate.pptx"
outputFile ="MergeSelectedSlides.pptx"
#Create an instance of presentation document
ppt = Presentation()
#Remove the first slide
ppt.Slides.RemoveAt(0)
#Load two PPT files
ppt1 = Presentation()
ppt1.LoadFromFile(inputFile_1)
ppt2 = Presentation()
ppt2.LoadFromFile(inputFile_2)
#Append all slides in ppt1 to ppt
for i, unusedItem in enumerate(ppt1.Slides):
    ppt.Slides.AppendBySlide(ppt1.Slides[i])
#Append the second slide in ppt2 to ppt
ppt.Slides.AppendBySlide(ppt2.Slides[1])
#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()
