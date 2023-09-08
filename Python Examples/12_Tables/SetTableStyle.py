from spire.presentation.common import *
from spire.presentation import *


inputFile = "./Data/SetTableStyle.pptx"
outputFile = "SetTableStyle.pptx"


#Creat a ppt document and load file
ppt = Presentation()
ppt.LoadFromFile(inputFile)

#Get tbe table
table = None
for shape in ppt.Slides[0].Shapes:
    if isinstance(shape, ITable):
        table = shape

        #Set the table style from TableStylePreset and apply it to selected table
        table.StylePreset = TableStylePreset.MediumStyle1Accent2
#Save the file
ppt.SaveToFile(outputFile, FileFormat.Pptx2010)
ppt.Dispose()



    
