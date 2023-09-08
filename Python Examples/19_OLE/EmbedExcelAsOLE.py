from spire.presentation.common import *
from spire.presentation import *

inputFile = "./Data/EmbedExcelAsOLE.xlsx"
outputFile = "EmbedExcelAsOLE.pptx"

# Create a Presentaion document
ppt = Presentation()

# Load a image file
stream = Stream("Data/EmbedExcelAsOLE.png")
oleImage = ppt.Images.AppendStream(stream)
stream.Close()

rec = RectangleF.FromLTRB(80, 60, oleImage.Width+80, oleImage.Height+60)
# Insert an OLE object to presentation based on the Excel data
oleStream = Stream(inputFile)
oleObject = ppt.Slides[0].Shapes.AppendOleObject("excel", oleStream, rec)
oleObject.SubstituteImagePictureFillFormat.Picture.EmbedImage = oleImage
oleObject.ProgId = "Excel.Sheet.12"
oleStream.Close()
# Save the document
ppt.SaveToFile(outputFile, FileFormat.Pptx2010)
ppt.Dispose()
