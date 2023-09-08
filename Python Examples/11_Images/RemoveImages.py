from spire.presentation.common import *
from spire.presentation import *

inputFile = "./Data/RemoveImages.pptx"
outputFile = "RemoveImages.pptx"


ppt = Presentation()
ppt.LoadFromFile(inputFile)

slide = ppt.Slides[0]

for i in range(slide.Shapes.Count - 1, -1, -1):
    #It is the SlidePicture object
    if isinstance(slide.Shapes[i], SlidePicture):
        slide.Shapes.RemoveAt(i)

#Save to file.
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()

    


    
