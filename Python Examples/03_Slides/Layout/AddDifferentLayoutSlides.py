from spire.presentation.common import *
from spire.presentation import *


outputFile ="AddDifferentLayoutSlides.pptx"

#Create a PPT document
presentation = Presentation()

#Remove the default slide
presentation.Slides.RemoveAt(0)

#Loop through slide layouts
for slideLayoutType in SlideLayoutType:
    #Append slide by specifing slide layout
    presentation.Slides.AppendByLayoutType(slideLayoutType)

#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()