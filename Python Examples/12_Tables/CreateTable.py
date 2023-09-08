from spire.presentation.common import *
import math
from spire.presentation import *


inputFile = "./Data/CreateTable.pptx"
outputFile = "CreateTable.pptx"


#Create a PPT document
presentation = Presentation()

#Load the document from disk
presentation.LoadFromFile(inputFile)

widths = [100, 100, 150, 100, 100]
heights = [15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15]

#Add new table to PPT
left = math.trunc(presentation.SlideSize.Size.Width / float(2)) - 275;
table = presentation.Slides[0].Shapes.AppendTable(left, 90, widths, heights)

dataStr = [["Name", "Capital", "Continent", "Area", "Population"], ["Venezuela", "Caracas", "South America", "912047", "19700000"], ["Bolivia", "La Paz", "South America", "1098575", "7300000"], ["Brazil", "Brasilia", "South America", "8511196", "150400000"], ["Canada", "Ottawa", "North America", "9976147", "26500000"], ["Chile", "Santiago", "South America", "756943", "13200000"], ["Colombia", "Bagota", "South America", "1138907", "33000000"], ["Cuba", "Havana", "North America", "114524", "10600000"], ["Ecuador", "Quito", "South America", "455502", "10600000"], ["Paraguay", "Asuncion", "South America", "406576", "4660000"], ["Peru", "Lima", "South America", "1285215", "21600000"], ["Jamaica", "Kingston", "North America", "11424", "2500000"], ["Mexico", "Mexico City", "North America", "1967180", "88600000"]]

#Add data to table
for i in range(0, 13):
    for j in range(0, 5):
        #Fill the table with data
        table[j,i].TextFrame.Text = dataStr[i][j]

        #Set the Font
        table[j,i].TextFrame.Paragraphs[0].TextRanges[0].LatinFont = TextFont("Arial Narrow")

#Set the alignment of the first row to Center
for i in range(0, 5):
    table[i,0].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Center

#Set the style of table
table.StylePreset = TableStylePreset.LightStyle3Accent1

#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()

        


    
