from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/BordersAndShading.pptx"
outputFile ="BordersAndShading.pptx"

#Load a PPT document
presentation = Presentation()
presentation.LoadFromFile(inputFile)

shape = presentation.Slides[0].Shapes[0]

#Set line color and width of the border
shape.Line.FillType = FillFormatType.Solid
shape.Line.Width = 3
shape.Line.SolidFillColor.Color = Color.get_LightYellow()

#Set the gradient fill color of shape
shape.Fill.FillType = FillFormatType.Gradient
shape.Fill.Gradient.GradientShape = GradientShapeType.Linear
shape.Fill.Gradient.GradientStops.AppendByKnownColors(1, KnownColors.LightBlue)
shape.Fill.Gradient.GradientStops.AppendByKnownColors(0, KnownColors.LightSkyBlue)

#Set the shadow for the shape
shadow = OuterShadowEffect()
shadow.BlurRadius = 20
shadow.Direction = 30
shadow.Distance = 8
shadow.ColorFormat.Color = Color.get_LightSeaGreen()
shape.EffectDag.OuterShadowEffect = shadow

#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2007)
presentation.Dispose()