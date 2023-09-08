from spire.presentation import *

inputFile = "./Data/SetBackground.pptx"
inputImg = "./Data/bg.png"
outputFile = "SetBackground_out.pptx"

#Create PPT document
presentation = Presentation()
presentation.LoadFromFile(inputFile)
#Set the background of the first slide to Gradient color
presentation.Slides[0].SlideBackground.Type = BackgroundType.Custom
presentation.Slides[0].SlideBackground.Fill.FillType = FillFormatType.Gradient
presentation.Slides[0].SlideBackground.Fill.Gradient.GradientShape = GradientShapeType.Linear
presentation.Slides[0].SlideBackground.Fill.Gradient.GradientStyle = GradientStyle.FromCorner1
presentation.Slides[0].SlideBackground.Fill.Gradient.GradientStops.AppendByKnownColors(1, KnownColors.SkyBlue)
presentation.Slides[0].SlideBackground.Fill.Gradient.GradientStops.AppendByKnownColors(0, KnownColors.White)
#Set the background of the second slide to Solid color
presentation.Slides[1].SlideBackground.Type = BackgroundType.Custom
presentation.Slides[1].SlideBackground.Fill.FillType = FillFormatType.Solid
presentation.Slides[1].SlideBackground.Fill.SolidColor.Color = Color.get_SkyBlue()
presentation.Slides.Append()
#Set the background of the third slide to picture
stream = Stream(inputImg)
imageData = presentation.Images.AppendStream(stream)
presentation.Slides[2].SlideBackground.Type = BackgroundType.Custom
presentation.Slides[2].SlideBackground.Fill.FillType = FillFormatType.Picture
presentation.Slides[2].SlideBackground.Fill.PictureFill.FillType = PictureFillType.Stretch
presentation.Slides[2].SlideBackground.Fill.PictureFill.Picture.EmbedImage = imageData
#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()