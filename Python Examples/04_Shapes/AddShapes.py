from spire.presentation.common import *
from spire.presentation import *

outputFile ="AddShapes.pptx"
#Create PPT document
presentation = Presentation()
#Set background Image
ImageFile = "Data/bg.png"
rect = RectangleF.FromLTRB (0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height)
presentation.Slides[0].Shapes.AppendEmbedImageByPath( ShapeType.Rectangle, ImageFile, rect)
presentation.Slides[0].Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.get_FloralWhite()
#Append new shape - Triangle and set style
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Triangle, RectangleF.FromLTRB (115, 130, 215, 230))
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_LightGreen()
shape.ShapeStyle.LineColor.Color = Color.get_White()
#Append new shape - Ellipse
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Ellipse, RectangleF.FromLTRB (290, 130, 440, 230))
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_LightSkyBlue()
shape.ShapeStyle.LineColor.Color = Color.get_White()
#Append new shape - Heart
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Heart, RectangleF.FromLTRB (470, 130, 600, 230))
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_Red()
shape.ShapeStyle.LineColor.Color = Color.get_LightGray()
#Append new shape - FivePointedStar
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.FivePointedStar, RectangleF.FromLTRB (90, 270, 240, 420))
shape.Fill.FillType = FillFormatType.Gradient
shape.Fill.SolidColor.Color = Color.get_Black()
shape.ShapeStyle.LineColor.Color = Color.get_White()
#Append new shape - Rectangle
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (320, 290, 420, 410))
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_Pink()
shape.ShapeStyle.LineColor.Color = Color.get_LightGray()
#Append new shape - BentUpArrow
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.BentUpArrow, RectangleF.FromLTRB (470, 300, 720, 400))
#Set the color of shape
shape.Fill.FillType = FillFormatType.Gradient
shape.Fill.Gradient.GradientStops.AppendByKnownColors(1, KnownColors.Olive)
shape.Fill.Gradient.GradientStops.AppendByKnownColors(0, KnownColors.PowderBlue)
shape.ShapeStyle.LineColor.Color = Color.get_Red()
#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()