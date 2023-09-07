from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/Animations.pptx"
outputFile ="Animations.pptx"
#Create a PPT document
presentation = Presentation()
#Load the document from disk
presentation.LoadFromFile(inputFile)
#Add title
rec_title = RectangleF.FromLTRB (50, 200, 250, 250)
shape_title = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, rec_title)
shape_title.ShapeStyle.LineColor.Color = Color.get_Transparent()
shape_title.Fill.FillType =FillFormatType.none
para_title = TextParagraph()
para_title.Text = "Animations:"
para_title.Alignment = TextAlignmentType.Center
para_title.TextRanges[0].LatinFont = TextFont("Myriad Pro Light")
para_title.TextRanges[0].FontHeight = 32
para_title.TextRanges[0].IsBold = TriState.TTrue
para_title.TextRanges[0].Fill.FillType = FillFormatType.Solid
para_title.TextRanges[0].Fill.SolidColor.Color = Color.FromArgb(255,68, 68, 68)
shape_title.TextFrame.Paragraphs.Append(para_title)
#Set the animation of slide to Circle
presentation.Slides[0].SlideShowTransition.Type = TransitionType.Circle
#Append new shape - Triangle
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Triangle, RectangleF.FromLTRB (100, 280, 180, 360))
#Set the color of shape
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_CadetBlue()
shape.ShapeStyle.LineColor.Color = Color.get_White()
#Set the animation of shape
shape.Slide.Timeline.MainSequence.AddEffect(shape, AnimationEffectType.Path4PointStar)
#Append new shape - Rectangle and set animation
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (210, 280, 360, 360))
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_CadetBlue()
shape.ShapeStyle.LineColor.Color = Color.get_White()
shape.AppendTextFrame("Animated Shape")
shape.Slide.Timeline.MainSequence.AddEffect(shape, AnimationEffectType.FadedSwivel)
#Append new shape - Cloud and set the animation
shape = presentation.Slides[0].Shapes.AppendShape(ShapeType.Cloud, RectangleF.FromLTRB (390, 280, 470, 360))
shape.Fill.FillType = FillFormatType.Solid
shape.Fill.SolidColor.Color = Color.get_White()
shape.ShapeStyle.LineColor.Color = Color.get_CadetBlue()
shape.Slide.Timeline.MainSequence.AddEffect(shape, AnimationEffectType.FadedZoom)
#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()
