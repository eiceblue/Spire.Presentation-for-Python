from spire.presentation.common import *
from spire.presentation import *


License.SetLicenseKey("");

inputFile ="././Data/SetTransitions.pptx"
outputFile ="SetTransitions.pptx"

#Create PPT document
presentation = Presentation()

#Load the PPT with password
presentation.LoadFromFile(inputFile)

#Set the first slide transition as push and sound mode
presentation.Slides[0].SlideShowTransition.Type = TransitionType.Push
presentation.Slides[0].SlideShowTransition.SoundMode = TransitionSoundMode.StartSound

#Set the second slide transition as circle and set the speed 
presentation.Slides[1].SlideShowTransition.Type = TransitionType.Fade
presentation.Slides[1].SlideShowTransition.Speed = TransitionSpeed.Slow

#Save the file
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()