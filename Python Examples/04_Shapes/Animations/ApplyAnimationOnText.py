from spire.presentation.common import *
from spire.presentation import *

outputFile ="ApplyAnimationOnText.pptx"
#Create an instance of presentation document
ppt = Presentation()
#Get the first slide
slide = ppt.Slides[0]
#Set background image
ImageFile = "./Data/bg.png"
rect = RectangleF.FromLTRB (0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height)
slide.Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, ImageFile, rect)
slide.Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.get_FloralWhite()
#Add a shape to the slide
shape = slide.Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (250, 150, 450, 250))
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_LightBlue()
shape.ShapeStyle.LineColor.Color = Color.get_White()
shape.AppendTextFrame("This demo shows how to apply animation on text in PPT document.")
#Apply animation to the text in shape
animation = shape.Slide.Timeline.MainSequence.AddEffect(shape, AnimationEffectType.Float)
animation.SetStartEndParagraphs(0, 0)
#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()