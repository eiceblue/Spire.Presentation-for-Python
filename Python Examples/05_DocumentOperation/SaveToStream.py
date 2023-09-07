from spire.presentation import *

inputFile = "./Data/bg.png"
outputFile = "SaveToStream.pptx"

#Create PowerPoint file and save it to stream
presentation = Presentation()
#Set background Image
rect = RectangleF.FromLTRB (0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height)
presentation.Slides[0].Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, inputFile, rect)
presentation.Slides[0].Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.get_FloralWhite()
#Append new shape
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (50, 100, 650, 250))
shape.Fill.FillType = FillFormatType.none
shape.ShapeStyle.LineColor.Color = Color.get_White()
#Add text to shape
shape.TextFrame.Text = "This demo shows how to Create PowerPoint file and save it to Stream."
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.FillType = FillFormatType.Solid
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.get_Black()
shape.TextFrame.Paragraphs[0].TextRanges[0].FontHeight = 30
#Save to Stream
stream = Stream(outputFile)
presentation.SaveToFile(stream, FileFormat.Pptx2013)
stream.Close()
presentation.Dispose()