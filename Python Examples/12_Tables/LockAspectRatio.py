from spire.presentation.common import *
from spire.presentation import *


inputFile = "./Data/Table.pptx"
outputFile = "LockAspectRatio.pptx"


#Create a PPT document
presentation = Presentation()

#Load PPT file from disk
presentation.LoadFromFile(inputFile)
#Get the first slide
slide = presentation.Slides[0]
for shape in slide.Shapes:
    #Verify if it is table
    if isinstance(shape, ITable):
        table = shape
        #Lock aspect ratio
        table.ShapeLocking.AspectRatioProtection = True

presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()


    
