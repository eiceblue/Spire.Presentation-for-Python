from spire.presentation.common import *
from spire.presentation import *
#pip install Pillow
from PIL import Image


inputFile = "./Data/BlankSample_N.pptx"
outputFile = "InsertEMFInPPT.pptx"


#Load a PPT document
presentation = Presentation()
presentation.LoadFromFile(inputFile)

#EMF file path
ImageFile = "./Data/InsertEMF.emf"

#Define image size
img = Image.open(ImageFile)
width = img.width / 1.5
height = img.height / 1.5
rect = RectangleF.FromLTRB (100, 100, width+100, height+100)

#Append the EMF in slide
image = presentation.Slides[0].Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, ImageFile, rect)
image.Line.FillType = FillFormatType.none

#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()



    
