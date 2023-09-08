from spire.presentation import *

inputFile = "./Data/PPTSample_N.pptx"
outputFile = "SetGradientBackground.pptx"

#Create a PPT document
presentation = Presentation()
#Load document from disk
presentation.LoadFromFile(inputFile)
#Get the first slide
slide = presentation.Slides[0]
#Set the background to gradient
slide.SlideBackground.Type = BackgroundType.Custom
slide.SlideBackground.Fill.FillType = FillFormatType.Gradient
#Add gradient stops
slide.SlideBackground.Fill.Gradient.GradientStops.AppendByColor(0.1, Color.get_LightSeaGreen())
slide.SlideBackground.Fill.Gradient.GradientStops.AppendByColor(0.7, Color.get_LightCyan())
#Set gradient shape type
slide.SlideBackground.Fill.Gradient.GradientShape = GradientShapeType.Linear
#Set the angle
slide.SlideBackground.Fill.Gradient.LinearGradientFill.Angle = 45
#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()