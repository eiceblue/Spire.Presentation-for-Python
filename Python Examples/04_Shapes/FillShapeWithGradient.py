from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/FillShapeWithGradient.pptx"
outputFile ="FillShapeWithGradient.pptx"

#Load a PPT document
ppt = Presentation()
ppt.LoadFromFile(inputFile)
#Get the first shape and set the style to be Gradient
GradientShape = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IAutoShape) else None
GradientShape.Fill.FillType = FillFormatType.Gradient
GradientShape.Fill.Gradient.GradientStops.AppendByColor (0, Color.get_LightSkyBlue())
GradientShape.Fill.Gradient.GradientStops.AppendByColor(1, Color.get_LightGray())
ppt.SaveToFile(outputFile, FileFormat.Pptx2010)
ppt.Dispose()