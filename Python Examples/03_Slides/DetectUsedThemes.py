from spire.presentation.common import *
from spire.presentation import *


inputFile ="./Data/Themes.pptx"
outputFile ="DetectUsedThemes.txt"
#Create an instance of presentation document
ppt = Presentation()
#Load file
ppt.LoadFromFile(inputFile)
sb = []
themeName = ""
sb.append ("This is the name list of the used theme below.")
#Get the theme name of each slide in the document
for slide in ppt.Slides:
    themeName = slide.Theme.Name
    sb.append (themeName)
#Save to the text document
fp = open(outputFile,"w")
for s in sb:
    fp.write(s + "\n")
fp.close()
ppt.Dispose()
