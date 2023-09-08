from spire.presentation.common import *
from spire.presentation import *


inputFile = "./Data/ExtractImage.pptx"
outputFile = "ChangeImageSize.pptx"


#Create a PPT document
presentation = Presentation()

#Load document from disk
presentation.LoadFromFile(inputFile)


scale = 0.5
for slide in presentation.Slides:
    for shape in slide.Shapes:
        if isinstance(shape, IEmbedImage):
            image = shape if isinstance(shape, IEmbedImage) else None
            image.Width = image.Width * scale
            image.Height = image.Height * scale

#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()

    


    
