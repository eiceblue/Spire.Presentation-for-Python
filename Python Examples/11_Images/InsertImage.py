from spire.presentation.common import *
import math
from spire.presentation import *


inputFile = "./Data/InsertImage.pptx"
outputFile = "InsertImage_out.pptx"


#Load a PPT document
presentation = Presentation()
presentation.LoadFromFile(inputFile)

#Insert image to PPT
ImageFile2 = "./Data/Logo1.png"
left = math.trunc(presentation.SlideSize.Size.Width / float(2)) - 280
rect1 = RectangleF.FromLTRB (left, 140, 120+left, 260)
image = presentation.Slides[0].Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, ImageFile2, rect1)
image.Line.FillType = FillFormatType.none

#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()

        

    
