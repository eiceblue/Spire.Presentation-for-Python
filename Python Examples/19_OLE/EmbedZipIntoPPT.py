from spire.presentation.common import *
from spire.presentation import *

inputFile = "./Data/EmbedZipIntoPPT.pptx"
inputFile_z = "./Data/test.zip"
inputFile_I = "./Data/icon.png"
outputFile = "EmbedZipIntoPPT.pptx"

# Create a Presentaion document
ppt = Presentation()
ppt.LoadFromFile(inputFile)

# Load a zip object
stream = Stream(inputFile_z)

rec = RectangleF.FromLTRB(80, 60, 180, 160)

# Insert the zip object to presentation
ole = ppt.Slides[0].Shapes.AppendOleObject(inputFile_z, stream, rec)
ole.ProgId = "Package"
image = Stream(inputFile_I)
oleImage = ppt.Images.AppendStream(image)
ole.SubstituteImagePictureFillFormat.Picture.EmbedImage = oleImage

# Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2010)
ppt.Dispose()
