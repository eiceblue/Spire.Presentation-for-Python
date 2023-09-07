from spire.presentation import *

inputFile = "./Data/InputTemplate.pptx"
outputFile = "SetShowTypeAsKiosk.pptx"

#Create an instance of presentation document
ppt = Presentation()
#Load file
ppt.LoadFromFile(inputFile)
#Specify the presentation show type as kiosk
ppt.ShowType = SlideShowType.Kiosk
#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()