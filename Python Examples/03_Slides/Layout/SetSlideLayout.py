from spire.presentation.common import *
from spire.presentation import *


outputFile ="SetSlideLayout.pptx"

#Create an instance of presentation document
ppt = Presentation()

#Remove the first slide
ppt.Slides.RemoveAt(0)

#Append a slide and set the layout for slide
slide = ppt.Slides.AppendByLayoutType(SlideLayoutType.Title)

#Add content for Title and Text
shape = slide.Shapes[0] if isinstance(slide.Shapes[0], IAutoShape) else None
shape.TextFrame.Text = "Hello Wolrd! -> This is title"

shape = slide.Shapes[1] if isinstance(slide.Shapes[1], IAutoShape) else None
shape.TextFrame.Text = "E-iceblue Support Team -> This is content"

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()