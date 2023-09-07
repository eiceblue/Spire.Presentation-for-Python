from spire.presentation.common import *
from spire.presentation import *


inputFile ="./Data/ExtractText.pptx"
outputFile ="ExtractText.txt"

#Create a PPT document and load file
presentation = Presentation()
presentation.LoadFromFile(inputFile)

sb = []
#Foreach the slide and extract text
for slide in presentation.Slides:
    for shape in slide.Shapes:
        if isinstance(shape, IAutoShape):
            for tp in ( shape if isinstance(shape, IAutoShape) else None).TextFrame.Paragraphs:
                sb.append (tp.Text)

fp = open(outputFile,"w",encoding = 'utf-8')
for s in sb:
    fp.write(s + "\n")
fp.close()
presentation.Dispose()