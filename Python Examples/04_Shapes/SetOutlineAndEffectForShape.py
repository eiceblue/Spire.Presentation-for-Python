from spire.presentation import *

inputFile = "./Data/bg.png"
outputFile = "SetOutlineAndEffectForShape.pptx"

#Create an instance of presentation document
ppt = Presentation()
#Get the first slide
slide = ppt.Slides[0]
#Set background Image
rect = RectangleF.FromLTRB (0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height)
slide.Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, inputFile, rect)
slide.Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.get_FloralWhite()
#Draw a Rectangle shape
shape = slide.Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (150, 180, 250, 230))
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_SkyBlue()
#Set outline color
shape.ShapeStyle.LineColor.Color = Color.get_Red()
#Set shadow effect
shadow = PresetShadow()
shadow.ColorFormat.Color = Color.get_LightSkyBlue()
shadow.Preset = PresetShadowValue.FrontRightPerspective
shadow.Distance = 10.0
shadow.Direction = 225.0
shape.EffectDag.PresetShadowEffect = shadow
#Draw a Ellipse shape
shape = slide.Shapes.AppendShape(ShapeType.Ellipse, RectangleF.FromLTRB (400, 150, 500, 250))
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_SkyBlue()
#Set outline color
shape.ShapeStyle.LineColor.Color = Color.get_Yellow()
#Set shadow effect
glow = GlowEffect()
glow.ColorFormat.Color = Color.get_LightPink()
glow.Radius = 20.0
shape.EffectDag.GlowEffect = glow
#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()