from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/Template_Az.pptx"
outputFile ="LineSpacing.pptx"

#Create a PPT document
presentation = Presentation()

#Load PPT file from disk
presentation.LoadFromFile(inputFile)
#Get the first slide
slide = presentation.Slides[0]
#Add a shape 
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (50, 100, presentation.SlideSize.Size.Width - 50, 400))
shape.ShapeStyle.LineColor.Color = Color.get_White()
shape.Fill.FillType = FillFormatType.none
shape.TextFrame.Paragraphs.Clear()

#Add text
shape.AppendTextFrame("Spire.Presentation for Python is a professional presentation processing API that is highly compatible with PowerPoint. It is a completely independent class library that developers can use to create, edit, convert, and save PowerPoint presentations efficiently without installing Microsoft PowerPoint.")
#Set font and color of text
textRange = shape.TextFrame.TextRange
textRange.Fill.FillType = FillFormatType.Solid
textRange.Fill.SolidColor.Color = Color.get_BlueViolet()
textRange.FontHeight = 20
textRange.LatinFont = TextFont("Lucida Sans Unicode")

#Set properties of paragraph
shape.TextFrame.Paragraphs[0].SpaceBefore = 100
shape.TextFrame.Paragraphs[0].SpaceAfter = 100
shape.TextFrame.Paragraphs[0].LineSpacing = 150

presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()