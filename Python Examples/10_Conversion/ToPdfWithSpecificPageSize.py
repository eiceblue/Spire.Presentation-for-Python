from spire.presentation.common import *
from spire.presentation import *


inputFile = "./Data/ToPDF.pptx"
outputFile = "ToPdfWithSpecificPageSize.pdf"


#Create a PPT document
presentation = Presentation()

#Load PPT file from disk
presentation.LoadFromFile(inputFile)

#Set A4 page size
presentation.SlideSize.Type = SlideSizeType.A4

#Set landscape orientation
presentation.SlideSize.Orientation = SlideOrienation.Landscape

#Save the PPT to PDF file format
presentation.SaveToFile(outputFile, FileFormat.PDF)
presentation.Dispose()

       


    
