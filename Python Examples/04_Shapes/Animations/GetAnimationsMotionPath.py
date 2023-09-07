from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/GetAnimationsMotionPath.pptx"
outputFile ="GetAnimationsMotionPath.txt"
presentation = Presentation()
presentation.LoadFromFile(inputFile)
slide = presentation.Slides[0]
#Get the first shape
shape = slide.Shapes[0]
#Create a StringBuilder to save the tracks
sb = []
i = 1
#Traverse all animations
for effect in shape.Slide.Timeline.MainSequence:
    if effect.ShapeTarget.Id==shape.Id:
        #Get MotionPath
        path = (effect.CommonBehaviorCollection[0]).Path        
        #Get all points in the path
        for motionCmdPath in path:
            points = motionCmdPath.Points
            comType = motionCmdPath.CommandType
            if points is not None:
                for point in points:
                    sb.append(str(i) + "  MotionType: " + str(comType )+ " -> X: " + str(point.X) + ", Y: " + str(point.Y))
                i += 1
fp = open(outputFile,"w")
for s in sb:
    fp.write(s + "\n")
presentation.Dispose()
