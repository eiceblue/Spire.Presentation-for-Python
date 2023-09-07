from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/Animation.pptx"
outputFile ="SetAnimationForAnimateText.pptx"
#Create an instance of presentation document
ppt = Presentation()
#Load file
ppt.LoadFromFile(inputFile)
#Set the AnimateType as Letter
ppt.Slides[0].Timeline.MainSequence[0].IterateType = AnimateType.Letter
#Set the IterateTimeValue for the animate text
ppt.Slides[0].Timeline.MainSequence[0].IterateTimeValue = 10
#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()
