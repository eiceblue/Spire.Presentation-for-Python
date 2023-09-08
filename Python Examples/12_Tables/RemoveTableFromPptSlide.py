from spire.presentation.common import *
from spire.presentation import *


inputFile = "./Data/Template_Ppt_1.pptx"
outputFile = "RemoveTableFromPptSlide.pptx"


#Create a PPT document
presentation = Presentation()

#Load the file from disk.
presentation.LoadFromFile(inputFile)

#Get the tables within the PPT document.
shape_tems = []

for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, ITable):
        #Add new table to table list.
        shape_tems.append(shape)

#Remove all the tables form the first slide.
for shape in shape_tems:
    presentation.Slides[0].Shapes.Remove(shape)

#Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()

     


    
