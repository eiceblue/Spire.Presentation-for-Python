from spire.presentation.common import *
from spire.presentation import *

inputFile = "./Data/audio.pptx"
outputFile = "HideAudioDuringShow.pptx"

# Load a PPT document
presentation = Presentation()
presentation.LoadFromFile(inputFile)

# Get the first slide
slide = presentation.Slides[0]

# Hide Audio during show
for shape in slide.Shapes:
    if isinstance(shape, IAudio):
        shape.HideAtShowing = True

# Save the file
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
