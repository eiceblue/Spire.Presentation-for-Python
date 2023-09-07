from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/Template_Az1.pptx"
outputFile ="GetTextFrameEffectiveData.txt"

#Create a PPT document
presentation = Presentation()

#Load PPT file from disk
presentation.LoadFromFile(inputFile)
#Get the first slide
slide = presentation.Slides[0]
#Get a shape 
shape = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IAutoShape) else None

textFrameFormat = shape.TextFrame
sb = []
sb.append ("Anchoring type: " + str(textFrameFormat.AnchoringType))
sb.append("Autofit type: " + str(textFrameFormat.AutofitType))
sb.append("Text vertical type: " + str(textFrameFormat.VerticalTextType))
sb.append("Margins")
sb.append("   Left: " + str(textFrameFormat.MarginLeft))
sb.append("   Top: " + str(textFrameFormat.MarginTop))
sb.append("   Right: " + str(textFrameFormat.MarginRight))
sb.append("   Bottom: " + str(textFrameFormat.MarginBottom))


fp = open(outputFile,"w",encoding = 'utf-8')
for s in sb:
    fp.write(s + "\n")
fp.close()

presentation.Dispose()