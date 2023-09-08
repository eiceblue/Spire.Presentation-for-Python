from spire.presentation.common import *
from spire.presentation import *


inputFile = "./Data/video.pptx"
inputFile2 = "./Data/repleaceVido.mp4"
outputFile = "ReplaceVideo.pptx"

# Create PPT document
presentation = Presentation()

# Load the PPT document from disk.
presentation.LoadFromFile(inputFile)

videos = presentation.Videos

# Traverse all the slides of PPT file
for sld in presentation.Slides:
    # Traverse all the shapes of slides
    for sp in sld.Shapes:
        # If shape is IVideo
        if isinstance(sp, IVideo):
            # Replace the video
            video = sp if isinstance(sp, IVideo) else None
            # Load the video document from disk.
            stream = Stream(inputFile2)
            videoData = videos.AppendByStream(stream)
            video.EmbeddedVideoData = videoData
# Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2016)
presentation.Dispose()
