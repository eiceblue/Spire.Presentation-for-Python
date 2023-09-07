from spire.presentation import *

inputFile = "./Data/bg.png"
outputFile = "PreventOrAllowChangingShape_out.pptx"

#Create an instance of presentation document
ppt = Presentation()
#Set background image
rect = RectangleF.FromLTRB (0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height)
ppt.Slides[0].Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, inputFile, rect)
ppt.Slides[0].Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.get_FloralWhite()
#Add a rectangle shape to the slide
shape = ppt.Slides[0].Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (50, 100, 450, 250))
#Set the shape format
shape.Fill.FillType = FillFormatType.none
shape.ShapeStyle.LineColor.Color = Color.get_LightBlue()
shape.TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Justify
shape.TextFrame.Text = "Demo for locking shapes:\n    Green/Black stands for editable.\n    Grey stands for non-editable."
shape.TextFrame.Paragraphs[0].TextRanges[0].LatinFont = TextFont("Arial Rounded MT Bold")
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.FillType = FillFormatType.Solid
shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.get_Black()
#The changes of selection and rotation are allowed
shape.Locking.RotationProtection = False
shape.Locking.SelectionProtection = False
#The changes of size, position, shape type, aspect ratio, text editing and ajust handles are not allowed 
shape.Locking.ResizeProtection = True
shape.Locking.PositionProtection = True
shape.Locking.ShapeTypeProtection = True
shape.Locking.AspectRatioProtection = True
shape.Locking.TextEditingProtection = True
shape.Locking.AdjustHandlesProtection = True
#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()
