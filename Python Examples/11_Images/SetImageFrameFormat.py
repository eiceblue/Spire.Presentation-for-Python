from spire.presentation.common import *
import math
from spire.presentation import *

outputFile = "SetImageFrameFormat.pptx"


#Create a PPT document
presentation = Presentation()

#Load an image
imageFile = "./Data/iceblueLogo.png"

stream = Stream(imageFile)
imageData = presentation.Images.AppendStream(stream)
stream.Close()

#Add the image in document
rect = RectangleF.FromLTRB (100, 100, math.trunc(imageData.Width / float(2))+100, math.trunc(imageData.Height / float(2)) +100)
pptImage = presentation.Slides[0].Shapes.AppendEmbedImageByImageData (ShapeType.Rectangle, imageData, rect)

#Set the formatting of the image frame
pptImage.Line.FillFormat.FillType = FillFormatType.Solid
pptImage.Line.FillFormat.SolidFillColor.Color = Color.get_LightBlue()
pptImage.Line.Width = 5
pptImage.Rotation = -45

#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()


    


    
