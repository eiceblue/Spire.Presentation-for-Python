from spire.presentation import *

inputFile = "./Data/bg.png"
outputFile = "Set3DEffectForShape_out.pptx"

#Create an instance of presentation document
ppt = Presentation()
#Set background image
rect = RectangleF.FromLTRB (0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height)
ppt.Slides[0].Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, inputFile, rect)
ppt.Slides[0].Shapes[0].Line.FillFormat.SolidFillColor.Color = Color.get_FloralWhite()
#Add shape1 and fill it with color
shape1 = ppt.Slides[0].Shapes.AppendShape(ShapeType.RoundCornerRectangle, RectangleF.FromLTRB (150, 150, 300, 300))
shape1.Fill.FillType = FillFormatType.Solid
shape1.Fill.SolidColor.KnownColor = KnownColors.SkyBlue
#Initialize a new instance of the 3-D class for shape1 and set its properties
effect1 = shape1.ThreeD.ShapeThreeD
effect1.PresetMaterial = PresetMaterialType.Powder
effect1.TopBevel.PresetType = BevelPresetType.ArtDeco
effect1.TopBevel.Height = 4
effect1.TopBevel.Width = 12
effect1.BevelColorMode = BevelColorType.Contour
effect1.ContourColor.KnownColor = KnownColors.LightBlue
effect1.ContourWidth = 3.5
#Add shape2 and fill it with color
shape2 = ppt.Slides[0].Shapes.AppendShape(ShapeType.Pentagon, RectangleF.FromLTRB (400, 150, 550, 300))
shape2.Fill.FillType = FillFormatType.Solid
shape2.Fill.SolidColor.KnownColor = KnownColors.LightGreen
#Initialize a new instance of the 3-D class for shape2 and set its properties
effect2 = shape2.ThreeD.ShapeThreeD
effect2.PresetMaterial = PresetMaterialType.SoftEdge
effect2.TopBevel.PresetType = BevelPresetType.SoftRound
effect2.TopBevel.Height = 12
effect2.TopBevel.Width = 12
effect2.BevelColorMode = BevelColorType.Contour
effect2.ContourColor.KnownColor = KnownColors.LawnGreen
effect2.ContourWidth = 5
#Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()