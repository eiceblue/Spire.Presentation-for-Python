from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/Animation.pptx"
outputFile ="SetAnimationRepeatType.pptx"

#Create an instance of presentation document
presentation = Presentation()
#Load file
presentation.LoadFromFile(inputFile)
#Get the first slide
slide = presentation.Slides[0]
animations = slide.Timeline.MainSequence
animations[0].Timing.AnimationRepeatType = AnimationRepeatType.UtilEndOfSlide
#Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()
