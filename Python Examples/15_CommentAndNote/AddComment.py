from spire.presentation.common import *
from spire.presentation import *

inputFile = "./Data/AddComment.pptx"
outputFile = "AddComment.pptx"

# Create a PPT document and load file
presentation = Presentation()
presentation.LoadFromFile(inputFile)

# Comment author
author = presentation.CommentAuthors.AddAuthor("E-iceblue", "comment:")

# Add comment
point = PointF.Empty()
point.X = 18
point.Y = 25
presentation.Slides[0].AddComment(
    author, "Add comment", point, DateTime.get_Now())

# Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2010)
presentation.Dispose()
