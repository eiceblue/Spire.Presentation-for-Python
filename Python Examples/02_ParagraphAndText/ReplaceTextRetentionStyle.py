from spire.presentation.common import *
from spire.presentation import *


inputFile ="./Data/SomePresentation.pptx"
outputFile ="ReplaceTextRetentionStyle.pptx"

ppt = Presentation()
ppt.LoadFromFile(inputFile)
ppt.Slides[0].ReplaceFirstText("use", "test", True)
ppt.Slides[1].ReplaceAllText("Spire", "new spire", True)
#Save to file.
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()
