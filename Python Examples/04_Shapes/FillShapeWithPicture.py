from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/FillShapeWithPicture.pptx"
outputFile ="FillShapeWithPicture.pptx"

#Load a PPT document
ppt = Presentation()
ppt.LoadFromFile(inputFile)
#Get the first shape and set the style to be Gradient
shape = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IAutoShape) else None
#Fill the shape with picture
picUrl = "./Data/backgroundImg.png"
shape.Fill.FillType = FillFormatType.Picture
shape.Fill.PictureFill.Picture.Url = picUrl
shape.Fill.PictureFill.FillType = PictureFillType.Stretch
ppt.SaveToFile(outputFile, FileFormat.Pptx2010)
ppt.Dispose()
