from spire.presentation.common import *
from spire.presentation import *


inputFile = "./Data/ToSVG.pptx"
outputFile = "Svg/"


#Create PPT document
presentation = Presentation()

#Load PPT file from disk
presentation.LoadFromFile(inputFile)

#Retain note when converting a PPT document to SVG files
presentation.IsNoteRetained = True

for index,slide in enumerate(presentation.Slides):
    fileName = outputFile + "ToSVG-"+str(index)+".svg"
    svgStream = slide.SaveToSVG()
    svgStream.Save(fileName)

presentation.Dispose()

     


    
