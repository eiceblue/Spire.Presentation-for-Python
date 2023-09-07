from spire.presentation.common import *
from spire.presentation import *


outputFile ="./Data/SetShadowEffect.pptx"

#Create an instance of presentation document
ppt = Presentation()

#Set background image
ImageFile = "./Data/bg.png"
rect = RectangleF.FromLTRB (0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height)
ppt.Slides[0].Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, ImageFile, rect)
ppt.Slides[0].Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.get_FloralWhite()

#Get reference of the slide
slide = ppt.Slides[0]

#Add a new rectangle shape to the first slide
shape = slide.Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (120, 100, 570, 300))
shape.Fill.FillType = FillFormatType.none

#Add the text to the shape and set the font for the text
shape.AppendTextFrame("Text shading on slides")
shape.TextFrame.Paragraphs[0].TextRanges[0].LatinFont = TextFont("Arial Black")
shape.TextFrame.Paragraphs[0].TextRanges[0].FontHeight = 21
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.FillType = FillFormatType.Solid
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.get_Black()

#//Add inner shadow and set all necessary parameters
#InnerShadowEffect Shadow = InnerShadowEffect()

#Add outer shadow and set all necessary parameters
Shadow = OuterShadowEffect()

Shadow.BlurRadius = 0
Shadow.Direction = 50
Shadow.Distance = 10
Shadow.ColorFormat.Color = Color.get_LightBlue()

#shape.TextFrame.TextRange.EffectDag.InnerShadowEffect = Shadow
shape.TextFrame.TextRange.EffectDag.OuterShadowEffect = Shadow

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()