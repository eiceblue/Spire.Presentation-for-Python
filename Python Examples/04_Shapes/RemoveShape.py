from spire.presentation import *

inputFile = "./Data/FindShapeByAltText.pptx"
outputFile = "RemoveShape_out.pptx"

#Create a PPT document
presentation = Presentation()
#Load doucment from disk
presentation.LoadFromFile(inputFile)
#Loop through slides
for i, unusedItem in enumerate(presentation.Slides):
    slide = presentation.Slides[i]
    #Loop through shapes
    j = 0
    while j < slide.Shapes.Count:
        shape = slide.Shapes[j]
        #Find the shapes whose alternative text contain "Shape"
        if shape.AlternativeText.find("Shape") != -1:
            slide.Shapes.Remove(shape)
            j -= 1
        j += 1
#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()