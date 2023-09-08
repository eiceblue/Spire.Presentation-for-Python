from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/SaveChartAsImage.pptx"
outputFile = "SaveChartAsImage.png"

#Create an instance of presentation document
ppt = Presentation()
#Load PPT file from disk
ppt.LoadFromFile(inputFile)

#Save chart as image in .png format
image = ppt.Slides[0].Shapes.SaveAsImage(0)

image.Save(outputFile)
image.Close()
ppt.Dispose()

