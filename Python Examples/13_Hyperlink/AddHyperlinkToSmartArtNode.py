from spire.presentation.common import *
from spire.presentation import *


inputFile = "./Data/SmartArtNode.pptx"
outputFile = "AddHyperlinkToSmartArtNode.pptx"


ppt = Presentation()
ppt.LoadFromFile(inputFile)

#Get the smartArt shape
sr = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], ISmartArt) else None
#Add hylerlinks to the nodes
node = sr.Nodes[0]
node.Click = ClickHyperlink(ppt.Slides[1])
node = sr.Nodes[1]
node.Click = ClickHyperlink(ppt.Slides[2])
node = sr.Nodes[2]
node.Click = ClickHyperlink(ppt.Slides[3])

#Save to file.
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()



    
