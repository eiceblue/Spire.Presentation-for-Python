from spire.presentation.common import *
from spire.presentation import *

inputFile = "./Data/Template_Ppt_1.pptx"
outputFile = "SetTableBorderStyle.pptx"


#Create a PPT document
presentation = Presentation()

#Load the file from disk.
presentation.LoadFromFile(inputFile)

#Find the table by looping through all the slides, and then set borders for it. 
for slide in presentation.Slides:
    for shape in slide.Shapes:
        if isinstance(shape, ITable):
            for row in shape.TableRows:
                for cell in row:
                    cell.BorderTop.FillType = FillFormatType.Solid
                    cell.BorderBottom.FillType = FillFormatType.Solid
                    cell.BorderLeft.FillType = FillFormatType.Solid
                    cell.BorderRight.FillType = FillFormatType.Solid

#Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()

        


    
