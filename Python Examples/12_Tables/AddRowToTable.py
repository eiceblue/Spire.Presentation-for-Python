from spire.presentation.common import *
from spire.presentation import *


inputFile = "./Data/Template_Ppt_1.pptx"
outputFile = "AddRowToTable.pptx"


#Create a PowerPoint document.
presentation = Presentation()

#Load the file from disk.
presentation.LoadFromFile(inputFile)

#Get the table within the PowerPoint document.
table = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], ITable) else None

#Get the first row.
row = table.TableRows[1]

#Clone the row and add it to the end of table.
table.TableRows.Append(row)
rowCount = table.TableRows.Count

#Get the last row.
lastRow = table.TableRows[rowCount - 1]

#Set new data of the first cell of last row.
lastRow[0].TextFrame.Text = " The first added cell"

#Set new data of the second cell of last row.
lastRow[1].TextFrame.Text = " The second added cell"

#Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()

  


    
