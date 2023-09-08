from spire.presentation.common import *
from spire.presentation import *


def AppendAllText(fname: str, text: List[str]):
    fp = open(fname, "w")
    for s in text:
        fp.write(s + "\n")
    fp.close()


inputFile = "./Data/Animation.pptx"
outputFile = "ObtainSoundEffect.txt"

# Create an instance of presentation document
ppt = Presentation()
# Load file
ppt.LoadFromFile(inputFile)

# Get the first slide
slide = ppt.Slides[0]

# Get the audio in a time node
audio = slide.Timeline.MainSequence[0].TimeNodeAudios[0]

# Get the properties of the audio, such as sound name, volume or detect if it's mute
text = []
text.append("SoundName: " + audio.SoundName)
text.append("Volume: " + str(audio.Volume))
text.append("IsMute: " + str(audio.IsMute))

# Save the properties of the audio to Text file
AppendAllText(outputFile, text)
ppt.Dispose()
