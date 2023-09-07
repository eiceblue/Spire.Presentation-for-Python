from spire.presentation import *

inputFile = "./Data/bg.png"
outputFile = "SetShadowEffectForShape.pptx"

#Create an instance of presentation document
ppt = Presentation()
slide = ppt.Slides[0]
#Set background image
rect = RectangleF.FromLTRB (0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height)
slide.Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, inputFile, rect)
slide.Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.get_FloralWhite()
#Add a shape to slide.
rect1 = RectangleF.FromLTRB (200, 150, 500, 270)
shape = slide.Shapes.AppendShape(ShapeType.Rectangle, rect1)
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_LightBlue()
shape.Line.FillType = FillFormatType.none
shape.TextFrame.Text = "This demo shows how to apply shadow effect to shape."
shape.TextFrame.TextRange.Fill.FillType = FillFormatType.Solid
shape.TextFrame.TextRange.Fill.SolidColor.Color = Color.get_Black()
#Create an inner shadow effect through InnerShadowEffect object. 
innerShadow = InnerShadowEffect()
innerShadow.BlurRadius = 20
innerShadow.Direction = 0
innerShadow.Distance = 0
innerShadow.ColorFormat.Color = Color.get_Black()
#Apply the shadow effect to shape.
shape.EffectDag.InnerShadowEffect = innerShadow
#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()