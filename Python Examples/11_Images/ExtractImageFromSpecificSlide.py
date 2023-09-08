from spire.presentation.common import *
from spire.presentation import *

inputFile = "./Data/Images.pptx"
outputFile = "ExtractImageFromSpecificSlide/"


#Create an instance of presentation document
ppt = Presentation()
#Load file
ppt.LoadFromFile(inputFile)

#Get the pictures on the second slide and save them to image file
i = 0
#Traverse all shapes in the second slide
for s in ppt.Slides[1].Shapes:
    #It is the SlidePicture object
    if isinstance(s, SlidePicture):
        #Save to image
        ps = s if isinstance(s, SlidePicture) else None
        ps.PictureFill.Picture.EmbedImage.Image.Save(outputFile+"SlidePic_"+str(i)+".png")
        i += 1
    #It is the PictureShape object
    if isinstance(s, PictureShape):
        #Save to image
        ps = s if isinstance(s, PictureShape) else None

        ps.EmbedImage.Image.Save(outputFile+"SlidePic_"+str(i)+".png")
        i += 1
ppt.Dispose()


    
