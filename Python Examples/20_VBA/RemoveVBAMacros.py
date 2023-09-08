from spire.presentation.common import *
from spire.presentation import *

inputFile = "./Data/Macros.ppt"
outputFile = "RemoveVBAMacros.ppt"

presentation = Presentation()

# Load PPT file from disk
presentation.LoadFromFile(inputFile)
# Remove macros
# Note, at present it only can work on macros in PPT file, has not supported for PPTM file yet.
presentation.DeleteMacros()
presentation.SaveToFile(outputFile, FileFormat.PPT)
presentation.Dispose()
