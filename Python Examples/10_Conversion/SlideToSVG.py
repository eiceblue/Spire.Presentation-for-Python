from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/PPTSample_N.pptx"
outputFile = "SlideToSVG.svg"

#Load document from disk
presentation = Presentation()
presentation.LoadFromFile(inputFile)

#Get the first slide
slide = presentation.Slides[0]

#Save the slide to SVG bytes
svgStream = slide.SaveToSVG()

#Write the bytes to file
svgStream.Save(outputFile)
svgStream.Dispose()

