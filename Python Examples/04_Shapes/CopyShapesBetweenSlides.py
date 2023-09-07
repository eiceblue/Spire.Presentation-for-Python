from spire.presentation.common import *
from spire.presentation import *


inputFile ="./Data/CopyShapesBetweenSlides.pptx"
outputFile ="CopyShapesBetweenSlides.pptx"
#Load the sample document
ppt = Presentation()
ppt.LoadFromFile(inputFile)
#Define the source slide and target slide
sourceSlide = ppt.Slides[0]
targetSlide = ppt.Slides[1]
#Copy the first shape from the source slide to the target slide
targetSlide.Shapes.AddShape(sourceSlide.Shapes[0])
#Save the document to file 
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()
