from spire.presentation.common import *
from spire.presentation import *


inputFile = "./Data/linkedSlide.pptx"
outputFile = "GetLinkedSlide.txt"


#Create Presentation
presentation = Presentation()

#Load ppt file
presentation.LoadFromFile(inputFile)
strB = []

#Get the second slide
slide = presentation.Slides[1]

#Get the first shape of the second slide
shape = slide.Shapes[0] if isinstance(slide.Shapes[0], IAutoShape) else None

#Get the linked slide index
if shape.Click.ActionType == HyperlinkActionType.GotoSlide:
    targetSlide = shape.Click.TargetSlide
    strB.append ("Linked slide number = " + str(targetSlide.SlideNumber))

#Save
with open(outputFile, 'w') as f:
    for item in strB:
        f.write("%s\n" % item)

presentation.Dispose()

    

    
