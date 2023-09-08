from spire.presentation.common import *
from spire.presentation import *


def AppendAllText(fname: str, text: List[str]):
    fp = open(fname, "w")
    for s in text:
        fp.write(s + "\n")
    fp.close()


inputFile = "./Data/Template_Ppt_1.pptx"
outputFile_px = "AddAndGetSpeakerNotes.pptx"
outputFile_txt = "AddAndGetSpeakerNotes.txt"

# Create a PowerPoint document.
presentation = Presentation()

# Load the file from disk.
presentation.LoadFromFile(inputFile)

# Get the first slide and in the PowerPoint document.
slide = presentation.Slides[0]

# Get the NotesSlide in the first slide,if there is no notes, we need to add it firstly.
ns = slide.NotesSlide
if ns is None:
    ns = slide.AddNotesSlide()

# Add the text string as the notes.
ns.NotesTextFrame.Text = "Speak notes added by Spire.Presentation"

content = []
content.append(
    "The speaker notes added by Spire.Presentation is: " + ns.NotesTextFrame.Text)

# Save to PowerPoint file.
presentation.SaveToFile(outputFile_px, FileFormat.Pptx2013)

# Get the speaker notes and save to txt file.
AppendAllText(outputFile_txt, content)
presentation.Dispose()
