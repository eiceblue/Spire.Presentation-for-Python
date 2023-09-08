from spire.presentation.common import *
from spire.presentation import *


inputFile = "./Data/Template_Ppt_1.pptx"
outputFile = "EditTableDataAndStyle.pptx"


#Create a PPT document
presentation = Presentation()

#Load the file from disk.
presentation.LoadFromFile(inputFile)

#Store the data used in replacement in string [].
strs = ["Germany", "Berlin", "Europe", "0152458", "20860000"]

table = None

#Get the table in PowerPoint document.
for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, ITable):
        table = shape

        #Change the style of table.
        table.StylePreset = TableStylePreset.LightStyle1Accent2

        for i, unusedItem in enumerate(table.ColumnsList):
            #Replace the data in cell.
            table[i,2].TextFrame.Text = strs[i]

            #Set the highlightcolor.
            table[i,2].TextFrame.TextRange.HighlightColor.Color = Color.get_BlueViolet()

#Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()

       


    
