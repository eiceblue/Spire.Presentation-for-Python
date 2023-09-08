from spire.presentation.common import *
from spire.presentation import *


inputFile = "Data/ChartSample2.pptx"
outputFile = "GetValuesAndUnitFromAxis.txt"

sb = []

#Create PPT document and load file
ppt = Presentation()
ppt.LoadFromFile(inputFile)

#Get chart on the first slide
Chart = ppt.Slides[0].Shapes[0] if isinstance(ppt.Slides[0].Shapes[0], IChart) else None

#Get unit from primary category axis
majorUnit = Chart.PrimaryCategoryAxis.MajorUnit
majorUnitScale = Chart.PrimaryCategoryAxis.MajorUnitScale

sb.append (str(majorUnit))
sb.append(str(majorUnitScale))


#Get values from primary value axis
minValue = Chart.PrimaryValueAxis.MinValue
maxValue = Chart.PrimaryValueAxis.MaxValue

sb.append(str(minValue))
sb.append(str(maxValue))

#Save the document
fp = open(outputFile,"w")
for s in sb:
    fp.write(s + "\n")
fp.close()
ppt.Dispose()
