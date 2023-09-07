from spire.presentation.common import *
from spire.presentation import *



inputFile ="./Data/PPTSample_N.pptx"
outputFile ="RemoveUnusedLayoutMaster.pptx"
ppt = Presentation()
ppt.LoadFromFile(inputFile)
#Create an array list
layouts = []
for i, unusedItem in enumerate(ppt.Slides):
    #Get the layout used by slide
    layout = ppt.Slides[i].Layout
    layouts.append(layout.SlideID)
#Loop through masters and layouts
for i, unusedItem in enumerate(ppt.Masters):
    masterlayouts = ppt.Masters[i].Layouts
    for j in range(masterlayouts.Count - 1, -1, -1):
        if not masterlayouts[j].SlideID in layouts:
            #Remove unused layout
            masterlayouts.RemoveMasterLayout(j)
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()
