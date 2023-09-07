from spire.presentation import *

inputFile = "./Data/ShapeTemplate.pptx"
outputFile = "ResetShapeSizeAndPosition.pptx"

#Create an instance of presentation document
ppt = Presentation()
#Load file
ppt.LoadFromFile(inputFile)
#Define the original slide size
currentHeight = ppt.SlideSize.Size.Height
currentWidth = ppt.SlideSize.Size.Width
#Change the slide size as A3
ppt.SlideSize.Type = SlideSizeType.A3
#Define the new slide size
newHeight = ppt.SlideSize.Size.Height
newWidth = ppt.SlideSize.Size.Width
#Define the ratio from the old and new slide size
ratioHeight = newHeight / currentHeight
ratioWidth = newWidth / currentWidth
#Reset the size and position of the shape on the slide
for slide in ppt.Slides:
    for shape in slide.Shapes:
        shape.Height = shape.Height * ratioHeight
        shape.Width = shape.Width * ratioWidth
        shape.Left = shape.Left * ratioHeight
        shape.Top = shape.Top * ratioWidth
#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()