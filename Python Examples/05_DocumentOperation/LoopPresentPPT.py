from spire.presentation import *

inputFile = "./Data/InputTemplate.pptx"
outputFile = "LoopPresentPPT.pptx"

#Create an instance of presentation document
ppt = Presentation()
#Load file
ppt.LoadFromFile(inputFile)
#Set the Boolean value of ShowLoop as true
ppt.ShowLoop = True
#Set the PowerPoint document to show animation and narration
ppt.ShowAnimation = True
ppt.ShowNarration = True
#Use slide transition timings to advance slide
ppt.UseTimings = True
#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()