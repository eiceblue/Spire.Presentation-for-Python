from spire.presentation.common import *
from spire.presentation import *


outputFile ="AddLineWithTwoPoints.pptx"
ppt = Presentation()
#Get the first slide
slide = ppt.Slides[0]
#Add line with two points
line = slide.Shapes.AppendShapeByPoint(ShapeType.Line, PointF(50.0, 50.0), PointF(150.0, 150.0))
line.ShapeStyle.LineColor.Color = Color.get_Red()
line = slide.Shapes.AppendShapeByPoint(ShapeType.Line, PointF(150.0, 150.0), PointF(250.0, 50.0))
line.ShapeStyle.LineColor.Color = Color.get_Blue()
#Save to file.
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()
