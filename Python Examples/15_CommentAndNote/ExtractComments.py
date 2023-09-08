from spire.presentation.common import *
from spire.presentation import *


def AppendAllText(fname: str, text: List[str]):
    fp = open(fname, "w")
    for s in text:
        fp.write(s + "\n")
    fp.close()


inputFile = "./Data/Template_Ppt_5.pptx"
outputFile = "ExtractComments.txt"

# Create a PowerPoint document.
presentation = Presentation()

# Load the file from disk.
presentation.LoadFromFile(inputFile)

cs = []

# Get all comments from the first slide.
comments = presentation.Slides[0].Comments

# Save the comments in txt file.
i = 0
while i < len(comments):
    cs.append(comments[i].Text + "\r\n")
    i += 1

# Save to file.
AppendAllText(outputFile, cs)
presentation.Dispose()
