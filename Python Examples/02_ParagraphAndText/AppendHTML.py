from spire.presentation.common import *
from spire.presentation import *


inputFile ="./Data/AppendHTML.pptx"
outputFile ="AppendHTML.pptx"
       

#Create a PPT document
ppt = Presentation()
ppt.LoadFromFile(inputFile)
#Add a shape 
shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (150, 100, 350, 300))

#Clear default paragraphs 
shape.TextFrame.Paragraphs.Clear()

code = "<html><body><p>This is a paragraph</p></body></html>"

#Append HTML, and generate a paragraph with default style in PPT document.
shape.TextFrame.Paragraphs.AddFromHtml(code)
codeColor = "<html><body><p style=\" color:black \">This is a paragraph</p></body></html>"
#Append HTML with black setting
shape.TextFrame.Paragraphs.AddFromHtml(codeColor)

#Add another shape
shape1 = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (350, 100, 550, 300))

#Clear default paragraph 
shape1.TextFrame.Paragraphs.Clear()

#Change the fill format of shape
shape1.Fill.FillType = FillFormatType.Solid
shape1.Fill.SolidColor.Color = Color.get_White()

#Append HTML
shape1.TextFrame.Paragraphs.AddFromHtml(code)
par = shape1.TextFrame.Paragraphs[0]
#Change the fill color for paragraph
for tr in par.TextRanges:
    tr.Fill.FillType = FillFormatType.Solid
    tr.Fill.SolidColor.Color = Color.get_Black()

ppt.SaveToFile(outputFile, FileFormat.Pptx2010)
ppt.Dispose()