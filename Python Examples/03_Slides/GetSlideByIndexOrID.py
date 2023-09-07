from spire.presentation.common import *
from spire.presentation import *



inputFile ="./Data/BlankSample_N.pptx"
outputFile ="GetSlideByIndexOrID.pptx"
#Create a PPT document
presentation = Presentation()
#Load document from disk
presentation.LoadFromFile(inputFile)
#Get slide by index 0
slide1 = presentation.Slides[0]
#Append a shape in the slide
shape1 = slide1.Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (100, 100, 300, 200))
#Add text in the shape
shape1.TextFrame.Text = "Get slide by index"
#Get slide by slide ID
slide2 = presentation.FindSlide(presentation.Slides[1].SlideID)
#Append a shape in the slide
shape2 = slide2.Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (100, 100, 300, 200))
#Add text in the shape
shape2.TextFrame.Text = "Get slide by slide id"
#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()
