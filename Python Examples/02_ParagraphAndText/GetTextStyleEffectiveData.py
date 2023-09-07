from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/Template_Az1.pptx"
outputFile ="GetTextStyleEffectiveData.txt"

#Create a PPT document
presentation = Presentation()

#Load PPT file from disk
presentation.LoadFromFile(inputFile)
#Get the first slide
slide = presentation.Slides[0]
#Get a shape 
shape = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IAutoShape) else None

sb =[]
for p, unusedItem in enumerate(shape.TextFrame.Paragraphs):
    paragraph = shape.TextFrame.Paragraphs[p]
    sb.append ("Text style for Paragraph " + str(p) + " :")
    #Get the paragraph style
    sb.append(" Indent: " + str(paragraph.Indent))
    sb.append(" Alignment: " + str(paragraph.Alignment))
    sb.append(" Font alignment: " + str(paragraph.FontAlignment))
    sb.append(" Hanging punctuation: " + str(paragraph.HangingPunctuation))
    sb.append(" Line spacing: " + str(paragraph.LineSpacing))
    sb.append(" Space before: " + str(paragraph.SpaceBefore))
    sb.append(" Space after: " + str(paragraph.SpaceAfter))
    sb.append("")
    for r, unusedItem in enumerate(paragraph.TextRanges):
        textRange = paragraph.TextRanges[r]
        sb.append("  Text style for Paragraph " + str(p) + " TextRange " + str(r) + " :")
        #Get the text range style
        sb.append("    Font height: " + str(textRange.FontHeight))
        sb.append("    Language: " + str(textRange.Language))
        sb.append("    Font: " + str(textRange.LatinFont.FontName))
        sb.append("")

fp = open(outputFile,"w",encoding = 'utf-8')
for s in sb:
    fp.write(s + "\n")
fp.close()

presentation.Dispose()

