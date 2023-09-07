from spire.presentation import *

inputFile = "./Data/InputTemplate.pptx"
outputFolder = "output"

#Create an instance of presentation document
ppt = Presentation()
#Load file
ppt.LoadFromFile(inputFile)
for i, slide in enumerate(ppt.Slides):
    #Initialize another instance of Presentation, and remove the blank slide
    newppt = Presentation()
    newppt.Slides.RemoveAt(0)
    #Append the specified slide from old presentation to the new one
    newppt.Slides.AppendBySlide(slide)
    #Save the document
    result =outputFolder + "//" + "SplitPPT-"+str(i)+".pptx"
    newppt.SaveToFile(result, FileFormat.Pptx2010)
    newppt.Dispose()
ppt.Dispose()