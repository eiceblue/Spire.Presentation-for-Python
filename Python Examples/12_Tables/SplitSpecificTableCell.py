from spire.presentation.common import *
from spire.presentation import *


inputFile = "./Data/Template_Ppt_1.pptx"
outputFile = "SplitSpecificTableCell.pptx"


#Create a PPT document
presentation = Presentation()

#Load the file from disk.
presentation.LoadFromFile(inputFile)

#Get the first slide.
slide = presentation.Slides[0]

#Get the table.
table = slide.Shapes[0]

#Split cell [1, 2] into 3 rows and 2 columns.
table[1,2].Split(3, 2)

#Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()

       


    
