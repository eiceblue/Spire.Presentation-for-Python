from spire.presentation.common import *
from spire.presentation import *


inputFile = "./Data/UpdateImage.pptx"
outputFile = "UpdateImage.pptx"


#Create an instance of presentation document
ppt = Presentation()
#Load file
ppt.LoadFromFile(inputFile)

#Get the first slide
slide = ppt.Slides[0]

#Append a new image to replace an existing image
stream = Stream("./Data/iceblueLogo.png")
image = ppt.Images.AppendStream(stream)
stream.Close()

#Replace the image which title is "image1" with the new image
for shape in slide.Shapes:
    if isinstance(shape, SlidePicture):
        if shape.AlternativeTitle == "image1":
            ( shape if isinstance(shape, SlidePicture) else None).PictureFill.Picture.EmbedImage = image

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()

       

    
