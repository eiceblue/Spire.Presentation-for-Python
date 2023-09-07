from spire.presentation.common import *
from spire.presentation import *


inputFile ="./Data/HideSlide.pptx"
outputFile ="HideSlide.pptx"
#Create a PPT document and load PPT file from disk
ppt = Presentation()
ppt.LoadFromFile(inputFile)
#Hide the second slide
ppt.Slides[1].Hidden = True
#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2010)
ppt.Dispose()