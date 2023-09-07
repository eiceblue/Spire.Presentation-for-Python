from spire.presentation.common import *
from spire.presentation import *


inputFile ="./Data/AppendSlideWithMasterLayout.pptx"
outputFile ="AppendSlideWithMasterLayout.pptx"
#Create a PPT document
presentation = Presentation()
#Load the document from disk
presentation.LoadFromFile(inputFile)
#Get the master
master = presentation.Masters[0]
#Get master layout slides
masterLayouts = master.Layouts
layoutSlide = masterLayouts[1]
#Append a rectangle to the layout slide
shape = layoutSlide.Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (10, 50, 110, 130))
#Add a text into the shape and set the style
shape.Fill.FillType = FillFormatType.none
shape.AppendTextFrame("Layout slide 1")
shape.TextFrame.Paragraphs[0].TextRanges[0].LatinFont = TextFont("Arial Black")
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.FillType = FillFormatType.Solid
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.get_CadetBlue()
#Append new slide with master layout
presentation.Slides.Append(presentation.Slides[0], master.Layouts[1])
#Another way to append new slide with master layout
presentation.Slides.Insert(2, presentation.Slides[1], master.Layouts[1])
#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()

