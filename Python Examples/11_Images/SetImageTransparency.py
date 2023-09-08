from spire.presentation.common import *
from spire.presentation import *

outputFile = "SetImageTransparency.pptx"


#Create an instance of presentation document
ppt = Presentation()

#Create an Image from the specified file
imagePath = "./Data/Logo1.png"

image = Image.FromFile(imagePath)
width = image.Width
height = image.Height
rect1 = RectangleF.FromLTRB (200, 100, width+200, height+100)
#Add a shape
shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rect1)
shape.Line.FillType = FillFormatType.none
#Fill shape with image
shape.Fill.FillType = FillFormatType.Picture
shape.Fill.PictureFill.Picture.Url = imagePath
shape.Fill.PictureFill.FillType = PictureFillType.Stretch
#Set transparency on image
shape.Fill.PictureFill.Picture.Transparency = 50

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()


       


    
