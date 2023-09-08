from spire.presentation.common import *
from spire.presentation import *

inputFile = "./Data/Template_Ppt_5.pptx"
outputFile = "SVG/PptToSvgRetainNotes/"

# Create a PowerPoint document.
presentation = Presentation()

# Load the file from disk.
presentation.LoadFromFile(inputFile)

# Retain the notes while converting PowerPoint file to svg file.
presentation.IsNoteRetained = True

# Convert presentation slides to svg file.
for index, slide in enumerate(presentation.Slides):
    stream = slide.SaveToSVG()
    stream.Save(outputFile+"output_"+str(index)+".svg")
    stream.Close()

presentation.Dispose()
