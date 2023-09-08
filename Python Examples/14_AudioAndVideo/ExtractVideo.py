from spire.presentation.common import *
from spire.presentation import *

inputFile = "./Data/video.pptx"
outputFile = "ExtractVideo/"

# Create PPT document
presentation = Presentation()

# Load the PPT document from disk.
presentation.LoadFromFile(inputFile)

# Define a variable
i = 0

# String for output file
result = outputFile + "ExtractVideo_"+str(i)+".avi"

# Traverse all the slides of PPT file
for slide in presentation.Slides:
    # Traverse all the shapes of slides
    for shape in slide.Shapes:
        # If shape is IVideo
        if isinstance(shape, IVideo):
            # Save the video
            shape.EmbeddedVideoData.SaveToFile(result)
            i += 1
presentation.Dispose()
