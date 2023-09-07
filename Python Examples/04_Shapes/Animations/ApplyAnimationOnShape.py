from spire.presentation.common import *
from spire.presentation import *

outputFile ="ApplyAnimationOnShape.pptx"
#Create an instance of presentation document
ppt = Presentation()
#Get the first slide
slide = ppt.Slides[0]
#Set background Image
ImageFile = "./Data/bg.png"
rect = RectangleF.FromLTRB (0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height)
slide.Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, ImageFile, rect)
slide.Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.get_FloralWhite()
#Insert a rectangle in the slide and fill the shape
shape = slide.Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (100, 150, 300, 230))
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_LightBlue()
shape.ShapeStyle.LineColor.Color = Color.get_White()
shape.AppendTextFrame("Animated Shape")
#Apply FadedSwivel animation effect to the shape
shape.Slide.Timeline.MainSequence.AddEffect(shape, AnimationEffectType.FadedSwivel)
#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()