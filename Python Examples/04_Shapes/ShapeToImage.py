from spire.presentation import *

inputFile = "./Data/ShapeToImage.pptx"
outputFolder = "output"

#Create a PPT document
presentation = Presentation()
presentation.LoadFromFile(inputFile)
for i, unusedItem in enumerate(presentation.Slides[0].Shapes):
    fileName =outputFolder + "//" + "ShapeToImage-"+str(i)+".png"
    #Save shapes as images
    image = presentation.Slides[0].Shapes.SaveAsImage(i)
    image.Save(fileName)
    image.Dispose()
presentation.Dispose()