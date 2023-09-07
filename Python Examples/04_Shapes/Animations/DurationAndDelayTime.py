from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/Animation.pptx"
outputFile ="DurationAndDelayTime.pptx"
#Create an instance of presentation document
presentation = Presentation()
presentation.LoadFromFile(inputFile)
#Get the first slide
slide = presentation.Slides[0]
animations = slide.Timeline.MainSequence
#Get duration time of animation
durationTime = animations[0].Timing.Duration
#Set new duration time of animation
animations[0].Timing.Duration = 0.8
#Get delay time of animation
delayTime = animations[0].Timing.TriggerDelayTime
#Set new delay time of animation
animations[0].Timing.TriggerDelayTime = 0.6
#Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()
