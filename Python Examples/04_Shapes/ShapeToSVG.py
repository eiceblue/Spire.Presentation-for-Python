from spire.presentation.common import *
from spire.presentation import *

inputFile = "./Demos/Data/ToSVG.pptx"
outputFile = "Svg/"

#Create PPT document
presentation = Presentation()
#Load PPT file from disk
presentation.LoadFromFile(inputFile)
# Counter for sequential SVG file names
m=0
# Iterate the slides in the presentation
for i in range(0,presentation.Slides.Count):
    slide = presentation.Slides[i]
    # Iterate the shapes in the slide
    for j in range(slide.Shapes.Count):
        shape = slide.Shapes[j]
        # Save the shape as SVG format
        stream = shape.SaveAsSvg()
        # Save the SVG stream to disk with a sequential filename
        stream.Save(outputFile + "SvgFile_" + str(m) + ".svg")
        # Flush and close the stream
        stream.Flush()
        stream.Close()
        # Increment the file counter
        m = m + 1
presentation.Dispose()