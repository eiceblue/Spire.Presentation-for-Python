from spire.presentation.common import *
from spire.presentation import *

inputFile = "./Data/Template_Ppt_1.pptx"
outputFile = "TraverseThroughCells.txt"


#Create a PowerPonit document.
presentation = Presentation()

#Load the file from disk.
presentation.LoadFromFile(inputFile)

content = []
content.append ("The data in cells of this PowerPoint file is: ")

#Get the table.
table = None
for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, ITable):
        table = shape

        #Traverse through the cells of table.
        for row in table.TableRows:
            for cell in row:
                content.append(cell.TextFrame.Text)

#Save to file.
with open(outputFile, 'w') as f:
    for item in content:
        f.write("%s\n" % item)
presentation.Dispose()

       


    
