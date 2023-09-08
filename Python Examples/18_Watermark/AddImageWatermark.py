from spire.presentation.common import *
from spire.presentation import *

inputFile = "Data/Template_Ppt_1.pptx"
outputFile = "output/AddImageWatermark.pptx"

# Create a PowerPoint document.
presentation = Presentation()

# Load the file from disk.
presentation.LoadFromFile(inputFile)

stream = Stream("Data/Logo.png")
image = presentation.Images.AppendStream(stream)
stream.Close()

# Set the properties of SlideBackground, and then fill the image as watermark.
presentation.Slides[0].SlideBackground.Type = BackgroundType.Custom
presentation.Slides[0].SlideBackground.Fill.FillType = FillFormatType.Picture
presentation.Slides[0].SlideBackground.Fill.PictureFill.FillType = PictureFillType.Stretch
presentation.Slides[0].SlideBackground.Fill.PictureFill.Picture.EmbedImage = image

# Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()
