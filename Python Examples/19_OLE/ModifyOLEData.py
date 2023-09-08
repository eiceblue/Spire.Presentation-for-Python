from spire.presentation.common import *
from spire.presentation import *

inputFile = "./Data/ModifyOLEData.pptx"
outputFile = "ModifyOLEData.pptx"

# Create a PPT document
presentation = Presentation()

# Load document from disk
presentation.LoadFromFile(inputFile)

# Loop through the slides and shapes
for slide in presentation.Slides:
    for shape in slide.Shapes:
        if isinstance(shape, IOleObject):
            # Find OLE object
            oleObject = shape if isinstance(shape, IOleObject) else None

            # Get its data and write to file
            stream = oleObject.Data
            stream2 = Stream()
            if oleObject.ProgId == "PowerPoint.Show.12":
                # Load the PPT stream
                ppt = Presentation()
                ppt.LoadFromStream(stream, FileFormat.Auto)
                # Append an image in slide
                ppt.Slides[0].Shapes.AppendEmbedImageByPath(
                    ShapeType.Rectangle, "Data/Logo.png", RectangleF.FromLTRB(50, 50, 150, 150))
                ppt.SaveToFile(stream2, FileFormat.Pptx2013)
                stream2.Position = 0
                # Modify the data
                oleObject.Data = stream2

# Save the document
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()
