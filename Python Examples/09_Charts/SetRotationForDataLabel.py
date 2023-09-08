from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/SetRotationForDataLabel.pptx"
outputFile = "SetRotationForDataLabel.pptx"
#Create an instance of presentation document
ppt = Presentation()
#Load file
ppt.LoadFromFile(inputFile)
Chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Set the rotation angle for the datalabels of first serie
for i, unusedItem in enumerate(Chart.Series[0].Values):
    datalabel = Chart.Series[0].DataLabels.Add()
    datalabel.ID = i
    datalabel.RotationAngle = 45
#Save to file.
ppt.SaveToFile(outputFile, FileFormat.Pptx2013)
ppt.Dispose()

