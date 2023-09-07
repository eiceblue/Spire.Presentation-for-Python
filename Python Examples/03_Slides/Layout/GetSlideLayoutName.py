from spire.presentation.common import *
from spire.presentation import *


inputFile ="././Data/GetSlideLayoutName.pptx"
outputFile ="GetSlideLayoutName.txt"

#Create a PPT document
presentation = Presentation()

#Load the document from disk
presentation.LoadFromFile(inputFile)

builder = []

#Loop through the slides of PPT document
for i, unusedItem in enumerate(presentation.Slides):
    #Get the name of slide layout
    name = presentation.Slides[i].Layout.Name
    builder.append ("The name of slide "+str(i)+" layout is: "+name)
       
     
fp = open(outputFile,"w",encoding = 'utf-8')
for s in builder:
    fp.write(s + "\n")
fp.close()

presentation.Dispose()