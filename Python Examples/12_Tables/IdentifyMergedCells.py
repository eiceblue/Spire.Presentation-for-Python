from spire.presentation.common import *
from spire.presentation import *

inputFile = "./Data/MergedCellInTable.pptx"
outputFile = "IdentifyMergedCells.txt"


#Create a PPT document
presentation = Presentation()

#Load PPT file from disk
presentation.LoadFromFile(inputFile)
#Get the first slide
slide = presentation.Slides[0]
strs = []
output = ""
for shape in slide.Shapes:
    #Verify if it is table
    if isinstance(shape, ITable):
        table = shape
        for r, unusedItem in enumerate(table.TableRows):
            for c, unusedItem in enumerate(table.ColumnsList):
                # Get cell
                currentCell = table.TableRows[r][c]
                #Identify if it is merged cell
                if currentCell.RowSpan > 1 or currentCell.ColSpan > 1:
                    output = "Cell {0:s}:{1:s} is a part of merged cell with RowSpan={2:s} and ColSpan={3:s} starting from Cell {4:s}:{5:s}.".format(str(r),str( c), str(currentCell.RowSpan), str(currentCell.ColSpan), str(currentCell.FirstRowIndex), str(currentCell.FirstColumnIndex))

                    strs.append (output)



with open(outputFile, 'w') as f:
    for item in strs:
        f.write("%s\n" % item)
presentation.Dispose()

    


    
