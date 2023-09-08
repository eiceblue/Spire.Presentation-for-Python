from spire.presentation.common import *
from spire.presentation import *


inputFile = "./Data/Template_Ppt_1.pptx"
outputFile = "SetBordersForExistingTable.pptx"


#Create a PPT document
presentation = Presentation()

#Load the file from disk.
presentation.LoadFromFile(inputFile)

#Get the table from the first slide of the sample document.
slide = presentation.Slides[0]
table = slide.Shapes[0] if isinstance(slide.Shapes[0], ITable) else None

#Set the border type as Inside and the border color as blue.
table.SetTableBorder(TableBorderType.Inside, 1, Color.get_Blue())

#Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()

   


    
