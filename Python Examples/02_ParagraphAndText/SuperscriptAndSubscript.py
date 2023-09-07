from spire.presentation.common import *
from spire.presentation import *


inputFile ="./Data/Template_Az.pptx"
outputFile ="SuperscriptAndSubscript.pptx"

#Create a PPT document
presentation = Presentation()

#Load PPT file from disk
presentation.LoadFromFile(inputFile)
#Get the first slide
slide = presentation.Slides[0]
#Add a shape 
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (150, 100, 350, 150))
shape.ShapeStyle.LineColor.Color = Color.get_White()
shape.Fill.FillType = FillFormatType.none
shape.TextFrame.Paragraphs.Clear()

shape.AppendTextFrame("Test")
tr = TextRange("superscript")
shape.TextFrame.Paragraphs[0].TextRanges.Append(tr)

#Set superscript text
shape.TextFrame.Paragraphs[0].TextRanges[1].Format.ScriptDistance = 30

textRange = shape.TextFrame.Paragraphs[0].TextRanges[0]
textRange.Fill.FillType = FillFormatType.Solid
textRange.Fill.SolidColor.Color =Color.get_Black()
textRange.FontHeight = 20
textRange.LatinFont = TextFont("Lucida Sans Unicode")

textRange = shape.TextFrame.Paragraphs[0].TextRanges[1]
textRange.Fill.FillType = FillFormatType.Solid
textRange.Fill.SolidColor.Color = Color.get_CadetBlue()
textRange.LatinFont = TextFont("Lucida Sans Unicode")


shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (150, 150, 350, 200))
shape.ShapeStyle.LineColor.Color = Color.get_White()
shape.Fill.FillType = FillFormatType.none
shape.TextFrame.Paragraphs.Clear()

shape.AppendTextFrame("Test")
tr = TextRange("subscript")
shape.TextFrame.Paragraphs[0].TextRanges.Append(tr)

#Set subscript text
shape.TextFrame.Paragraphs[0].TextRanges[1].Format.ScriptDistance = -25

textRange = shape.TextFrame.Paragraphs[0].TextRanges[0]
textRange.Fill.FillType = FillFormatType.Solid
textRange.Fill.SolidColor.Color = Color.get_Black()
textRange.FontHeight = 20
textRange.LatinFont = TextFont("Lucida Sans Unicode")

textRange = shape.TextFrame.Paragraphs[0].TextRanges[1]
textRange.Fill.FillType = FillFormatType.Solid
textRange.Fill.SolidColor.Color = Color.get_CadetBlue()
textRange.LatinFont = TextFont("Lucida Sans Unicode")


presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()