from spire.presentation import *

inputFile = "./Data/bg.png"
outputFile = "GroupShapes_out.pptx"

#Create an instance of presentation document
ppt = Presentation()
#Get the first slide
slide = ppt.Slides[0]
#Set background image
rect = RectangleF.FromLTRB (0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height)
slide.Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, inputFile, rect)
slide.Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.get_FloralWhite()
#Create two shapes in the slide
rectangle = slide.Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (250, 180, 450, 220))
rectangle.Fill.FillType = FillFormatType.Solid
rectangle.Fill.SolidColor.KnownColor = KnownColors.SkyBlue
rectangle.Line.Width = 0.1
ribbon = slide.Shapes.AppendShape(ShapeType.Ribbon2, RectangleF.FromLTRB (290, 155, 410, 235))
ribbon.Fill.FillType = FillFormatType.Solid
ribbon.Fill.SolidColor.KnownColor = KnownColors.LightPink
ribbon.Line.Width = 0.1
#Add the two shape objects to an array list
arr = []
arr.append(rectangle)
arr.append(ribbon)
#Group the shapes in the list
ppt.Slides[0].GroupShapes(arr)
#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()