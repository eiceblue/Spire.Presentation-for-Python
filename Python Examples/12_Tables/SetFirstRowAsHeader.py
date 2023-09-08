from spire.presentation.common import *
from spire.presentation import *


inputFile = "./Data/NormalTable.pptx"
outputFile = "SetFirstRowAsHeader.pptx"

table = None

#Load a PPT document
presentation = Presentation()
presentation.LoadFromFile(inputFile)

for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, ITable):
        table = shape

table.FirstRow = True

#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()

        

    
