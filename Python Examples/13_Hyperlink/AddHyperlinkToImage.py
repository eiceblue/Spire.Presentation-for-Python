from spire.presentation.common import *
from spire.presentation import *


inputFile = "./Data/Template_Ppt_5.pptx"
outputFile = "AddHyperlinkToImage.pptx"


#Create a PowerPoint document.
presentation = Presentation()

#Load the file from disk.
presentation.LoadFromFile(inputFile)

#Get the first slide.
slide = presentation.Slides[0]

#Add image to slide.
rect = RectangleF.FromLTRB (480, 350, 640, 510)
image = slide.Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, "./Data/Logo1.png", rect)

#Add hyperlink to the image.
hyperlink = ClickHyperlink("https://www.e-iceblue.com")
image.Click = hyperlink

#Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()

       

    
