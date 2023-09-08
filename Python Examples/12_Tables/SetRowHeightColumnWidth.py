from spire.presentation.common import *
from spire.presentation import *


inputFile = "./Data/SetRowHeightColumnWidth.pptx"
outputFile = "SetRowHeightColumnWidth.pptx"


#Creat a ppt document and load file
ppt = Presentation()
ppt.LoadFromFile(inputFile)

#Get the table
table = None
for shape in ppt.Slides[0].Shapes:
    if isinstance(shape, ITable):
        table = shape

        #Set the height for the rows
        table.TableRows[0].Height = 100
        table.TableRows[1].Height = 80
        table.TableRows[2].Height = 60
        table.TableRows[3].Height = 40
        table.TableRows[4].Height = 20

        #Set the column width
        table.ColumnsList[0].Width = 60
        table.ColumnsList[1].Width = 80
        table.ColumnsList[2].Width = 120
        table.ColumnsList[3].Width = 140
        table.ColumnsList[4].Width = 160
#Save the file
ppt.SaveToFile(outputFile, FileFormat.Pptx2010)
ppt.Dispose()


    
