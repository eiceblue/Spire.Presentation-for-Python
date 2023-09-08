from spire.presentation.common import *
from spire.presentation import *


inputFile = "./Data/ExtractImage.pptx"
outputFile = "ExtractImage/"


#Load a PPT document
ppt = Presentation()
ppt.LoadFromFile(inputFile)

for i, image in enumerate(ppt.Images):
    ImageName = outputFile+"Images_"+str(i)+".png"
    image.Image.Save(ImageName)

ppt.Dispose()

        

    
