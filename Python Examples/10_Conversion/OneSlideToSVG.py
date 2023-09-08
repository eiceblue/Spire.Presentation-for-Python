from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/OneSlideToSVG.pptx"
outputFile = "OneSlideToSVG.svg"

#Create PPT document
presentation = Presentation()

#Load PPT file from disk
presentation.LoadFromFile(inputFile)

#Convert the second slide to SVG
svgStream = presentation.Slides[1].SaveToSVG()
svgStream.Save(outputFile)
presentation.Dispose()

