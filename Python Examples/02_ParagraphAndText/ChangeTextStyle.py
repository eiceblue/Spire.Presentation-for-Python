from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/ChangeTextStyle.pptx"
outputFile ="ChangeTextStyle.pptx"

#Load a PPT document
presentation = Presentation()
presentation.LoadFromFile(inputFile)

shape = presentation.Slides[0].Shapes[0]
paras = shape.TextFrame.Paragraphs

#Set the style for the text content in the first paragraph
for tr in paras[0].TextRanges:
    tr.Fill.FillType = FillFormatType.Solid
    tr.Fill.SolidColor.Color = Color.get_ForestGreen()
    tr.LatinFont = TextFont("Lucida Sans Unicode")
    tr.FontHeight = 14

#Set the style for the text content in the third paragraph
for tr in paras[2].TextRanges:
    tr.Fill.FillType = FillFormatType.Solid
    tr.Fill.SolidColor.Color = Color.get_CornflowerBlue()
    tr.LatinFont = TextFont("Calibri")
    tr.FontHeight = 16
    tr.TextUnderlineType = TextUnderlineType.Dashed

#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2007)
presentation.Dispose()