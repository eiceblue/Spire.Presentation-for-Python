from spire.presentation import *
import math

inputFile = "./Data/Template_Ppt_7.pptx"
outputFile = "ResetPositionOfDateTimeAndSlideNumber.pptx"

#Create a PowerPoint document.
presentation = Presentation()
#Load the file from disk.
presentation.LoadFromFile(inputFile)
#Get the first slide from the sample document.
slide = presentation.Slides[0]
for shapeToMove in slide.Shapes:
    #Reset the position of the slide number to the left.
    if shapeToMove.Name.find ("Slide Number Placeholder") != -1:
        shapeToMove.Left = 0
    elif shapeToMove.Name.find ("Date Placeholder") != -1:
        #Reset the position of the date time to the center.
        shapeToMove.Left = math.trunc(presentation.SlideSize.Size.Width / float(2))
        #Reset the date time display style.
        ( shapeToMove if isinstance(shapeToMove, IAutoShape) else None).TextFrame.TextRange.Paragraph.Text = DateTime.get_Now().ToString("dd.MM.yyyy")
        ( shapeToMove if isinstance(shapeToMove, IAutoShape) else None).TextFrame.IsCentered = True
#Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()