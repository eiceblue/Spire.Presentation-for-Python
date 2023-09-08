from spire.presentation.common import *
from spire.presentation import *


inputFile = "./Data/GetBorderColorOfCell.pptx"
outputFile = "GetBorderColorOfCell.txt"

ppt = Presentation()
ppt.LoadFromFile(inputFile)

#Get the table in the first slide
table = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], ITable) else None

#Get borders' color of the first cell
sb = []
sb.append ("Color of left border:" + table[0,0].BorderLeftDisplayColor.ToString())
sb.append("Color of top border:" + table[0,0].BorderTopDisplayColor.ToString())
sb.append("Color of right border:" + table[0,0].BorderRightDisplayColor.ToString())
sb.append("Color of bottom border:" + table[0,0].BorderBottomDisplayColor.ToString())

#Get display color of the first cell
sb.append("Color of cell:" + table[0,0].DisplayColor.ToString())

with open(outputFile, 'w') as f:
    for item in sb:
        f.write("%s\n" % item)

ppt.Dispose()



    
