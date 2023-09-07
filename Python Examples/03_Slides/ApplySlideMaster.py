
from spire.presentation.common import *
from spire.presentation import *



inputFile ="./Data/InputTemplate.pptx"
outputFile ="ApplySlideMaster.pptx"

#Create an instance of presentation document
ppt = Presentation()
#Load file
ppt.LoadFromFile(inputFile)
#Get the first slide master from the presentation
masterSlide = ppt.Masters[0]
#Customize the background of the slide master
backgroundPic ="./Data/bg.png"
rect = RectangleF.FromLTRB (0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height)
masterSlide.SlideBackground.Fill.FillType = FillFormatType.Picture
image = masterSlide.Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, backgroundPic, rect)
masterSlide.SlideBackground.Fill.PictureFill.Picture.EmbedImage = image.PictureFill.Picture.EmbedImage
#Change the color scheme
masterSlide.Theme.ColorScheme.Accent1.Color = Color.get_Red()
masterSlide.Theme.ColorScheme.Accent2.Color = Color.get_RosyBrown()
masterSlide.Theme.ColorScheme.Accent3.Color = Color.get_Ivory()
masterSlide.Theme.ColorScheme.Accent4.Color = Color.get_Lavender()
masterSlide.Theme.ColorScheme.Accent5.Color = Color.get_Black()
#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()
