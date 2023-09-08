from spire.presentation.common import *
from spire.presentation import *

inputFile = ".\\Data\\audio.pptx"
outputFile = "ExtractAudio.wav"
AudioData = None

# Load a PPT document
presentation = Presentation()
presentation.LoadFromFile(inputFile)

for shape in presentation.Slides[0].Shapes:
    if isinstance(shape, IAudio):
        audio = shape if isinstance(shape, IAudio) else None
        AudioData = audio.Data
        AudioData.SaveToFile(outputFile)

presentation.Dispose()
