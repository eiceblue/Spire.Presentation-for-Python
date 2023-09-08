from spire.presentation.common import *
from spire.presentation import *


inputFile = "./Data/RemoveRowsAndColumns.pptx"
outputFile = "RemoveRowsAndColumns.pptx"


#Create a PPT document
presentation = Presentation()
presentation.LoadFromFile(inputFile)

#Get the table in PPT document
table = None
for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, ITable):
        table = shape

        #Remove the second column
        table.ColumnsList.RemoveAt(1, False)

        #Remove the second row
        table.TableRows.RemoveAt(1, False)
#Save and launch the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()



    
