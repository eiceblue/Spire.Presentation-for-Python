from spire.presentation.common import *
from spire.presentation import *


inputFile = "./Data/AddImageInTableCell.pptx"
outputFile = "AddImageInTableCell.pptx"


#Load a PPT document
ppt = Presentation()
ppt.LoadFromFile(inputFile)

#Get the first shape
table = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], ITable) else None

stream = Stream("./Data/Logo1.png")
pptImg = ppt.Images.AppendStream(stream)
stream.Close()

table[1,1].FillFormat.FillType = FillFormatType.Picture
table[1,1].FillFormat.PictureFill.Picture.EmbedImage = pptImg
table[1,1].FillFormat.PictureFill.FillType = PictureFillType.Stretch

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2010)
ppt.Dispose()

      


    
