from spire.presentation.common import *
from spire.presentation import *


inputFile ="./Data/TextBoxTemplate.pptx"
outputFile ="RemoveTextBox.pptx"

#Create an instance of presentation document
ppt = Presentation()
#Load file
ppt.LoadFromFile(inputFile)

#Get the first slide
slide = ppt.Slides[0]
#Traverse all the shapes in slide
i = 0
while i < slide.Shapes.Count:
    #Remove all shapes
    shape = slide.Shapes[i] if isinstance(slide.Shapes[i], IAutoShape) else None
    slide.Shapes.Remove(shape)

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()
