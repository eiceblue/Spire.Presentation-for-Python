from spire.presentation import *

inputFile = "./Data/FindShapeByAltText.pptx"
outputFile = "HideShape_out.pptx"

#Create a PPT document
presentation = Presentation()
#Load document from disk
presentation.LoadFromFile(inputFile)
#Loop through slides
for slide in presentation.Slides:
    #Loop through shapes in the slide
    for shape in slide.Shapes:
        #Find the shape whose alternative text is Shape1
        if shape.AlternativeText=="Shape1":
            #Hide the shape
            shape.IsHidden = True
#Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()