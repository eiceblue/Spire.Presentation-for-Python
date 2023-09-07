from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/Bullets.pptx"
outputFile ="Bullets.pptx"

#Load a PPT document
presentation = Presentation()
presentation.LoadFromFile(inputFile)

shape = presentation.Slides[0].Shapes[1]

for para in shape.TextFrame.Paragraphs:
    #Add the bullets
    para.BulletType = TextBulletType.Numbered
    para.BulletStyle = NumberedBulletStyle.BulletRomanLCPeriod


#Save the document and launch
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()