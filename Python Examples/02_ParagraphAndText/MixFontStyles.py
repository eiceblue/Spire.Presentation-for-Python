from spire.presentation.common import *
from spire.presentation import *
import re

inputFile ="./Data/FontStyle.pptx"
outputFile ="MixFontStyles.pptx"

#Create an instance of presentation document
ppt = Presentation()
#Load file
ppt.LoadFromFile(inputFile)

#Get the second shape of the first slide
shape = ppt.Slides[0].Shapes[1] if isinstance(ppt.Slides[0].Shapes[1], IAutoShape) else None
#Get the text from the shape 
originalText = shape.TextFrame.Text

#Split the string by specified words and return substrings to a string array
#splitArray = originalText.split(["bold", "red", "underlined", "bigger font size"])
splitArray = originalText.split("bold")

#Remove the paragraph from TextRange
tp = shape.TextFrame.Paragraphs[0]
tp.TextRanges.Clear()

#Append normal text that is in front of 'bold' to the paragraph
tr = TextRange(splitArray[0])
tp.TextRanges.Append(tr)
#Set font style of the text 'bold' as bold
tr = TextRange("bold")
tr.IsBold = TriState.TTrue
tp.TextRanges.Append(tr)

splitArray = splitArray[1].split("red")
#Append normal text that is in front of 'red' to the paragraph
tr = TextRange(splitArray[0])
tp.TextRanges.Append(tr)
#Set the color of the text 'red' as red
tr = TextRange("red")
tr.Fill.FillType = FillFormatType.Solid
tr.Format.Fill.SolidColor.Color = Color.get_Red()
tp.TextRanges.Append(tr)
splitArray = splitArray[1].split("underlined")
#Append normal text that is in front of 'underlined' to the paragraph
tr = TextRange(splitArray[0])
tp.TextRanges.Append(tr)
#Underline the text 'undelined'
tr = TextRange("underlined")
tr.TextUnderlineType = TextUnderlineType.Single
tp.TextRanges.Append(tr)

splitArray = splitArray[1].split("bigger font size")
#Append normal text that is in front of 'bigger font size' to the paragraph
tr = TextRange(splitArray[0])
tp.TextRanges.Append(tr)
#Set a large font for the text 'bigger font size'
tr = TextRange("bigger font size")
tr.FontHeight = 35
tp.TextRanges.Append(tr)

#Append other normal text
tr = TextRange(splitArray[1])
tp.TextRanges.Append(tr)

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()
