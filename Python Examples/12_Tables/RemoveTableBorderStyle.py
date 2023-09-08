from spire.presentation.common import *
from spire.presentation import *


inputFile = "./Data/Template_Ppt_1.pptx"
outputFile = "RemoveTableBorderStyle.pptx"


#Create a PowerPoint document.
presentation = Presentation()

#Load the file from disk.
presentation.LoadFromFile(inputFile)

for slide in presentation.Slides:
    for shape in slide.Shapes:
        if isinstance(shape, ITable):
            for row in shape.TableRows:
                for cell in row:
                    cell.BorderTop.FillType = FillFormatType.none
                    cell.BorderBottom.FillType = FillFormatType.none
                    cell.BorderLeft.FillType = FillFormatType.none
                    cell.BorderRight.FillType = FillFormatType.none

#Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()



    
