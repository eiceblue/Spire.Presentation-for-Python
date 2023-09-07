from spire.presentation.common import *
from spire.presentation import *


inputFile ="./Data/Titles.pptx"
outputFile ="GetAllTitles.txt"
#Create an instance of presentation document
ppt = Presentation()
#Load file
ppt.LoadFromFile(inputFile)
#Instantiate a list of IShape objects
shapelist = []
#Loop through all sildes and all shapes on each slide
for slide in ppt.Slides:
    for shape in slide.Shapes:
        if shape.Placeholder is not None:
            #Get all titles
            if shape.Placeholder.Type == PlaceholderType.Title:
                shapelist.append(shape)
            elif shape.Placeholder.Type == PlaceholderType.CenteredTitle:
                shapelist.append(shape)
            elif shape.Placeholder.Type == PlaceholderType.Subtitle:
                shapelist.append(shape)
#Loop through the list and get the inner text of all shapes in the list
sb = []
sb.append("Below are all the obtained titles:")
for i, unusedItem in enumerate(shapelist):
    shape1 = shapelist[i] if isinstance(shapelist[i], IAutoShape) else None
    sb.append (shape1.TextFrame.Text)
#Save to the Text file
fp = open(outputFile,"w")
for s in sb:
    fp.write(s + "\n")
fp.close()
ppt.Dispose()