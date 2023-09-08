from spire.presentation.common import *
from spire.presentation import *


inputFile = "./Data/Conversion.pptx"
outputFile = "ToSpecificSizeImage.png"


#Create an instance of presentation document
ppt = Presentation()
#Load file
ppt.LoadFromFile(inputFile)

#Save the first slide to Image and set the image size to 600*400
img = ppt.Slides[0].SaveAsImageByWH(600, 400)
#Save image to file
img.Save(outputFile)
img.Dispose()
ppt.Dispose()

      

    
