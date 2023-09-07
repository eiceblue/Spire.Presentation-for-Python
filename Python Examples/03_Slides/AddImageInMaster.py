from spire.presentation.common import *
from spire.presentation import *

# from spire.common import *
License.SetLicenseKey("")

inputFile ="./Data/AddImageInMaster.pptx"
outputFile ="AddImageInMaster.pptx"

#Create a PPT document
presentation = Presentation()

#Load the document from disk
presentation.LoadFromFile(inputFile)

#Get the master collection
master = presentation.Masters[0]

#Append image to slide master
image = "./Data/Logo.png"
rff = RectangleF.FromLTRB (40, 40, 130, 130)
pic = master.Shapes.AppendEmbedImageByPath(ShapeType.Rectangle, image, rff)
pic.Line.FillFormat.FillType = FillFormatType.none

#Add new slide to presentation
presentation.Slides.Append()

#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()
