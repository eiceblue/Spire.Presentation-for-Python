from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/Animation.pptx"
outputFile ="GetAnimationEffectInfo.txt"
presentation = Presentation()
presentation.LoadFromFile(inputFile)
sb = []
#Travel each slide
for slide in presentation.Slides:
    for effect in slide.Timeline.MainSequence:
        #Get the animation effect type
        animationEffectType = effect.AnimationEffectType
        sb.append ("animation effect type:" + str(animationEffectType))
        #Get the slide number where the animation is located
        slideNumber = slide.SlideNumber
        sb.append ("slide number:" + str(slideNumber))
        #Get the shape name
        shapeName = effect.ShapeTarget.Name
        sb.append ("shape name:" + shapeName + "\n")
fp = open(outputFile,"w")
for s in sb:
    fp.write(s + "\n")
presentation.Dispose()