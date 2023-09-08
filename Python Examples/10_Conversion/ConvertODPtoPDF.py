from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/toPdf.odp"
outputFile = "ConvertODPtoPDF.pdf"

#Create PPT document
presentation = Presentation()

#Load the PPT document from disk.
presentation.LoadFromFile(inputFile, FileFormat.ODP)

presentation.SaveToFile(outputFile,FileFormat.PDF)
presentation.Dispose()


