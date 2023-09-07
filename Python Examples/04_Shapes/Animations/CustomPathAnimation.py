from spire.presentation.common import *
from spire.presentation import *

outputFile ="CustomPathAnimation.pptx"
#Create PPT document
ppt = Presentation()
#Add shape
shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB(0, 0, 200, 200))
#Add animation
effect = ppt.Slides[0].Timeline.MainSequence.AddEffect(shape, AnimationEffectType.PathUser)
common = effect.CommonBehaviorCollection
motion = common[0]
motion.Origin = AnimationMotionOrigin.Layout
motion.PathEditMode = AnimationMotionPathEditMode.Relative
#Add moin path
moinPath = MotionPath()
p1=PointF(0.0,0.0)
p2=PointF(0.1,0.1)
p3=PointF(-0.1,0.2)
moinPath.Add(MotionCommandPathType.MoveTo, [p1], MotionPathPointsType.CurveAuto, True)
moinPath.Add(MotionCommandPathType.LineTo, [p2], MotionPathPointsType.CurveAuto, True)
moinPath.Add(MotionCommandPathType.LineTo, [p3], MotionPathPointsType.CurveAuto, True)
moinPath.Add(MotionCommandPathType.End, [], MotionPathPointsType.CurveStraight, True)
motion.Path = moinPath
#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2010)
ppt.Dispose()