from spire.presentation import *

inputFile = "./Data/GroupShapes.pptx"
outputFile = "UngroupShapes.pptx"

ppt = Presentation()
ppt.LoadFromFile(inputFile)
groupShape = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], GroupShape) else None
#Ungroup the shapes
ppt.Slides[0].Ungroup(groupShape)
#Save to file.
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()