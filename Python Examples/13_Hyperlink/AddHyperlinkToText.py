from spire.presentation.common import *
from spire.presentation import *

inputFile = "./Data/AddHyperlinkToText.pptx"
outputFile = "AddHyperlinkToText_out.pptx"


#Create a PowerPoint document.
presentation = Presentation()

#Load the file from disk.
presentation.LoadFromFile(inputFile)

#Find the text we want to add link to it.
shape = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IAutoShape) else None
tp = shape.TextFrame.Paragraphs[0]
temp = tp.Text

#Split the original text.
textToLink = "Spire.Presentation"
strSplit = temp.split("Spire.Presentation")

#Clear all text.
tp.TextRanges.Clear()

#Add new text.
tr = TextRange(strSplit[0])
tp.TextRanges.Append(tr)

#Add the hyperlink.
tr = TextRange(textToLink)
tr.ClickAction.Address = "https://www.e-iceblue.com/Introduce/presentation-for-python.html"
tp.TextRanges.Append(tr)

#Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()

    


    
