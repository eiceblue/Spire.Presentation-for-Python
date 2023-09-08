from spire.presentation.common import *
from spire.presentation import *


inputFile = "./Data/Template_Ppt_1.pptx"
outputFile = "FillParticularRowWithColor.pptx"


#Create a PPT document
presentation = Presentation()

#Load the file from disk.
presentation.LoadFromFile(inputFile)

#Fill particular table row with color.
table = None
for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, ITable):
        table = shape

        row = table.TableRows[1]
        for cell in row:
            cell.FillFormat.FillType = FillFormatType.Solid
            cell.FillFormat.SolidColor.Color = Color.get_Pink()

#Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()



    
