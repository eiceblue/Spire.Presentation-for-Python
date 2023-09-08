from spire.presentation.common import *
from spire.presentation import *

inputFile = "./Data/MergeTableCell.pptx"
outputFile = "MergeTableCell.pptx"


#Create a PPT document and load file
presentation = Presentation()
presentation.LoadFromFile(inputFile)

table = None
for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, ITable):
        table = shape

        #Merge the second row and third row of the first column
        table.MergeCells(table[0,1], table[0,2], False)

        table.MergeCells(table[3,4], table[4,4], True)


#Save and launch the file
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()



    
