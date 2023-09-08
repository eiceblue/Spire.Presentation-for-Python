from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/Template_Ppt_2.pptx"
outputFile = "ProtectChart.pptx"


#Create a PowerPonit document.
presentation = Presentation()

#Load the file from disk.
presentation.LoadFromFile(inputFile)

#Get the first shape from slide and convert it as IChart.
chart = presentation.Slides[0].Shapes[0] if isinstance(presentation.Slides[0].Shapes[0], IChart) else None

#Set the Boolean value of IChart.IsDataProtect as true.
chart.IsDataProtect = True

#Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()

    
        