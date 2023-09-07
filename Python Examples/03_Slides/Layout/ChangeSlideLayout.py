from spire.presentation.common import *
from spire.presentation import *

inputFile ="././Data/ChangeSlideLayout.pptx"
outputFile ="ChangeSlideLayout.pptx"


#Create a PPT document
presentation = Presentation()

#Load the document from disk
presentation.LoadFromFile(inputFile)

#Change the layout of slide
presentation.Slides[1].Layout = presentation.Masters[0].Layouts[4]

#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()