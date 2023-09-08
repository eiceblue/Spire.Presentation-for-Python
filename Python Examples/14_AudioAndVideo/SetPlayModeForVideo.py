from spire.presentation.common import *
from spire.presentation import *


inputFile = "./Data/Template_Ppt_8.pptx"
outputFile = "SetPlayModeForVideo.pptx"

# Create a PowerPoint document.
presentation = Presentation()

# Load the file from disk.
presentation.LoadFromFile(inputFile)

# Find the video by looping through all the slides and set its play mode as auto.
for slide in presentation.Slides:
    for shape in slide.Shapes:
        if isinstance(shape, IVideo):
            (shape if isinstance(shape, IVideo)
             else None).PlayMode = VideoPlayMode.Auto

# Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()
