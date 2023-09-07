from spire.presentation.common import *
from spire.presentation import *

outputFile ="CreateSlideMasterAndApply.pptx"
#Create an instance of presentation document
ppt = Presentation()
ppt.SlideSize.Type = SlideSizeType.Screen16x9
#Add slides
for i in range(0, 4):
    ppt.Slides.Append()
#Get the first default slide master
first_master = ppt.Masters[0]
#Append another slide master
ppt.Masters.AppendSlide(first_master)
second_master = ppt.Masters[1]
#Set different background image for the two slide masters
pic1 = "Data/bg.png"
pic2 = "Data/Setbackground.png"
#The first slide master
rect = RectangleF.FromLTRB (0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height)
first_master.SlideBackground.Fill.FillType = FillFormatType.Picture
image1 = first_master.Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, pic1, rect)
first_master.SlideBackground.Fill.PictureFill.Picture.EmbedImage = image1.PictureFill.Picture.EmbedImage
#The second slide master
second_master.SlideBackground.Fill.FillType = FillFormatType.Picture
image2 = second_master.Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, pic2, rect)
second_master.SlideBackground.Fill.PictureFill.Picture.EmbedImage = image2.PictureFill.Picture.EmbedImage
#Apply the first master with layout to the first slide
ppt.Slides[0].Layout = first_master.Layouts[1]
#Apply the second master with layout to other slides
for i in range(1, ppt.Slides.Count):
    ppt.Slides[i].Layout = second_master.Layouts[8]
#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()