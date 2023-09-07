from spire.presentation.common import *
from spire.presentation import *


inputFile ="././Data/SetTransitions.pptx"
outputFile ="SetTransitionEffects.pptx"


#Create PPT document
presentation = Presentation()

#Load the PPT
presentation.LoadFromFile(inputFile)

# Set effects
presentation.Slides[0].SlideShowTransition.Type = TransitionType.Cut
presentation.Slides[0].SlideShowTransition.Value.FromBlack = True

#Save the file
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()