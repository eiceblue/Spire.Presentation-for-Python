from spire.presentation.common import *
from spire.presentation import *


inputFile ="././Data/SetTransitions.pptx"
outputFile ="BetterSlideTransitions.pptx"

#Create PPT document
presentation = Presentation()

#Load the PPT
presentation.LoadFromFile(inputFile)

#Set the first slide transition as circle
presentation.Slides[0].SlideShowTransition.Type = TransitionType.Circle

# Set the transition time of 3 seconds
presentation.Slides[0].SlideShowTransition.AdvanceOnClick = True
presentation.Slides[0].SlideShowTransition.AdvanceAfterTime = 3000

#Set the second slide transition as comb and set the speed 
presentation.Slides[1].SlideShowTransition.Type = TransitionType.Comb
presentation.Slides[1].SlideShowTransition.Speed = TransitionSpeed.Slow

# Set the transition time of 5 seconds
presentation.Slides[1].SlideShowTransition.AdvanceOnClick = True
presentation.Slides[1].SlideShowTransition.AdvanceAfterTime = 5000

# Set the third slide transition as zoom
presentation.Slides[2].SlideShowTransition.Type = TransitionType.Zoom

# Set the transition time of 7 seconds
presentation.Slides[2].SlideShowTransition.AdvanceOnClick = True
presentation.Slides[2].SlideShowTransition.AdvanceAfterTime = 7000


#Save the file
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()