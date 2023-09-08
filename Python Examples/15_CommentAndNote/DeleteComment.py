from spire.presentation.common import *
from spire.presentation import *

inputFile = "./Data/DeleteComment.pptx"
outputFile = "DeleteComment.pptx"

# Create a PPT document and load file
presentation = Presentation()
presentation.LoadFromFile(inputFile)

# Replace the text in the comment
presentation.Slides[0].Comments[1].Text = "Replace comment"

# Delete the third comment
presentation.Slides[0].DeleteComment(presentation.Slides[0].Comments[2])

# Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()
