from spire.presentation.common import *
from spire.presentation import *

InputImageFile = "./Data/bg.png"
outputFile ="AutoFitTextOrShape.pptx"

#Create an instance of presentation document
ppt = Presentation()

#Set background image
       
rect = RectangleF.FromLTRB (0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height)
ppt.Slides[0].Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, InputImageFile, rect)
ppt.Slides[0].Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.get_FloralWhite()

#Set the AutofitType property to Shape
textShape2 = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (150, 100, 300, 180))
textShape2.TextFrame.Text = "Resize shape to fit text."
textShape2.TextFrame.AutofitType = TextAutofitType.Shape

#Set the AutofitType property to Normal
textShape1 = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (400, 100, 550, 180))
textShape1.TextFrame.Text = "Shrink text to fit shape. Shrink text to fit shape. Shrink text to fit shape. Shrink text to fit shape."
textShape1.TextFrame.AutofitType = TextAutofitType.Normal

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()