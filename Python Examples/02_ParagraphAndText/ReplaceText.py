from spire.presentation.common import *
from spire.presentation import *

def ReplaceTags(pSlide, TagValues):
    for curShape in pSlide.Shapes:
        if isinstance(curShape, IAutoShape):
            for tp in ( curShape if isinstance(curShape, IAutoShape) else None).TextFrame.Paragraphs:
                for curKey in TagValues.keys():
                    tp.Text = tp.Text.replace(curKey, TagValues[curKey])



inputFile ="./Data/TextTemplate.pptx"
outputFile ="ReplaceText.pptx"

tagValues = {}
tagValues["Spire.Presentation"] = "Spire.PPT"

#Create an instance of presentation document
ppt = Presentation()
#Load file
ppt.LoadFromFile(inputFile)

ReplaceTags(ppt.Slides[0], tagValues)

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()
 

