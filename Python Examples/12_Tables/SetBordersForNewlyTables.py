from spire.presentation.common import *
from spire.presentation import *

outputFile = "SetBordersForNewlyTables.pptx"


#Create a PPT document
presentation = Presentation()

#Set the table width and height for each table cell.
tableWidth = [100, 100, 100, 100, 100]
tableHeight = [20, 20]

#Traverse all the border type of the table.
for item in TableBorderType:
    #Add a table to the presentation slide with the setting width and height
    itable = presentation.Slides.Append().Shapes.AppendTable(100, 100, tableWidth, tableHeight)

    #Add some text to the table cell.
    itable.TableRows[0][0].TextFrame.Text = "Row"
    itable.TableRows[1][0].TextFrame.Text = "Column"

    #Set the border type, border width and the border color for the table.
    itable.SetTableBorder(item, 1.5, Color.get_Red())

#Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()

    


    
