from spire.presentation.common import *
from spire.presentation import *


outputFile ="Set3DEffectForText.pptx"

#Create a new presentation object
ppt = Presentation()

#Get the first slide
slide = ppt.Slides[0]

#Append a new shape to slide and set the line color and fill type
shape = slide.Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (30, 40, 680, 240))
shape.ShapeStyle.LineColor.Color = Color.get_White()
shape.Fill.FillType = FillFormatType.none

#Add text to the shape
shape.AppendTextFrame("This demo shows how to add 3D effect text to Presentation slide")

#Set the color of text in shape
textRange = shape.TextFrame.TextRange
textRange.Fill.FillType = FillFormatType.Solid
textRange.Fill.SolidColor.Color = Color.get_LightBlue()

#Set the Font of text in shape
textRange.FontHeight = 40
textRange.LatinFont = TextFont("Gulim")

#Set 3D effect for text
shape.TextFrame.TextThreeD.ShapeThreeD.PresetMaterial = PresetMaterialType.Matte
shape.TextFrame.TextThreeD.LightRig.PresetType = PresetLightRigType.Sunrise
shape.TextFrame.TextThreeD.ShapeThreeD.TopBevel.PresetType = BevelPresetType.Circle
shape.TextFrame.TextThreeD.ShapeThreeD.ContourColor.Color = Color.get_Green()
shape.TextFrame.TextThreeD.ShapeThreeD.ContourWidth = 3

#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()
