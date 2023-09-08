from spire.presentation.common import *
from spire.presentation import *


inputFile = "./Data/SetAlignmentInTable.pptx"
outputFile = "SetAlignmentInTable.pptx"


#Create a PPT document
presentation = Presentation()
presentation.LoadFromFile(inputFile)

table = None
for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, ITable):
        table = shape

        #Horizontal Alignment
        #Set the horizontal alignment for the cells in first column 
        table[0,1].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Left
        table[0,2].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Center
        table[0,3].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Right
        table[0,4].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Justify

        #Vertical Alignment
        #Set the vertical alignment for the cells in second column 
        table[1,1].TextAnchorType = TextAnchorType.Top
        table[1,2].TextAnchorType = TextAnchorType.Center
        table[1,3].TextAnchorType = TextAnchorType.Bottom
        table[1,4].TextAnchorType = TextAnchorType.none

        #Both orientaions
        #Set the both horizontal and vertical alignment for the cells in the third column 
        table[2,1].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Left
        table[2,1].TextAnchorType = TextAnchorType.Top

        table[2,2].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Right
        table[2,2].TextAnchorType = TextAnchorType.Center

        table[2,3].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Justify
        table[2,3].TextAnchorType = TextAnchorType.Bottom

        table[2,4].TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Center
        table[2,4].TextAnchorType = TextAnchorType.Top

#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()

    

    
