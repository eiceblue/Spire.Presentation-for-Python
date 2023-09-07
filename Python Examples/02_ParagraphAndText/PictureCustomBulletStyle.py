from spire.presentation.common import *
from spire.presentation import *


inputFile ="./Data/Bullets.pptx"
outputFile ="PictureCustomBulletStyle.pptx"
   
#Create an instance of presentation document
ppt = Presentation()
#Load file
ppt.LoadFromFile(inputFile)

#Get the second shape on the first slide
shape = ppt.Slides[0].Shapes[1] if isinstance(ppt.Slides[0].Shapes[1], IAutoShape) else None

#Traverse through the paragraphs in the shape
for paragraph in shape.TextFrame.Paragraphs:
    #Set the bullet style of paragraph as picture
    paragraph.BulletType = TextBulletType.Picture
    #Load a picture

    fileStream = Stream("./Data/icon.png")
   
    paragraph.BulletPicture.EmbedImage = ppt.Images.AppendStream (fileStream)
   
    fileStream.Close()
#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()