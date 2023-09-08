from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/ToImage.pptx"

#Create PPT document
presentation = Presentation()

#Load PPT file from disk
presentation.LoadFromFile(inputFile)

#Save PPT document to images
for i, slide in enumerate(presentation.Slides):
    fileName ="ToImage_img_"+str(i)+".png"
    image = slide.SaveAsImage()
    image.Save(fileName)
    image.Dispose()

presentation.Dispose()
