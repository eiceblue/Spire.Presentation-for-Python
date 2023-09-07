from spire.presentation.common import *
from spire.presentation import *



inputFile_1 = "./Data/CloneMaster1.pptx"
inputFile_2 = "./Data/CloneMaster2.pptx"
outputFile ="ClonePPTMasterToAnother.pptx"
#Load PPT1 from disk
presentation1 = Presentation()
presentation1.LoadFromFile(inputFile_1)
#Load PPT2 from disk
presentation2 = Presentation()
presentation2.LoadFromFile(inputFile_2)
#Add masters from PPT1 to PPT2
for masterSlide in presentation1.Masters:
    presentation2.Masters.AppendSlide(masterSlide)
#Save the document
presentation2.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation2.Dispose()