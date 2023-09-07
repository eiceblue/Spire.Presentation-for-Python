from spire.presentation.common import *
from spire.presentation import *

outputFile ="AddExitAnimationForShape.pptx"
#Create an instance of presentation document
ppt = Presentation()
#Get the first slide
slide = ppt.Slides[0]
#Set background image
ImageFile = "./Data/bg.png"
rect = RectangleF.FromLTRB (0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height)
slide.Shapes.AppendEmbedImageByPath(ShapeType.Rectangle, ImageFile, rect)
slide.Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.get_FloralWhite()
#Add a shape to the slide
starShape = slide.Shapes.AppendShape(ShapeType.FivePointedStar, RectangleF.FromLTRB (250, 100, 450, 300))
starShape.Fill.FillType = FillFormatType.Solid
starShape.Fill.SolidColor.KnownColor = KnownColors.LightBlue
#Add random bars effect to the shape
effect = slide.Timeline.MainSequence.AddEffect(starShape, AnimationEffectType.RandomBars)
#Change effect type from entrance to exit
effect.PresetClassType = TimeNodePresetClassType.Exit
#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()
