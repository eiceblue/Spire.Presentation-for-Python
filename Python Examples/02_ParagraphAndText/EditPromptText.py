from spire.presentation.common import *
from spire.presentation import *

inputFile ="./Data/HasPromptText.pptx"
outputFile ="EditPromptText.pptx"
#Load a PPT document
presentation = Presentation()
presentation.LoadFromFile(inputFile)

# Iterate through the slide
for shape in presentation.Slides[0].Shapes:
    if shape.Placeholder is not None and isinstance(shape, IAutoShape):
        text = ""
        # Set the text of the title
        if shape.Placeholder.Type == PlaceholderType.CenteredTitle:
            text = "custom title create by Spire"
        # Set text of the subtitle.
        elif shape.Placeholder.Type == PlaceholderType.Subtitle:
            text = "custom subtitle create by Spire"

        ( shape if isinstance(shape, IAutoShape) else None).TextFrame.Text = text
        
#Save the file
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()