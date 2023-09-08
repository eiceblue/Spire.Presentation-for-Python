from spire.presentation.common import *
from spire.presentation import *


def AppendAllText(fname: str, text: List[str]):
    fp = open(fname, "w")
    for s in text:
        fp.write(s + "\n")
    fp.close()


inputFile = "Data/ExtractTextFromSmartArt.pptx"
outputFile = "output/ExtractTextFromSmartArt.txt"

# Create a PowerPoint document.
presentation = Presentation()

# Load the file from disk.
presentation.LoadFromFile(inputFile)

# Traverse through all the slides of the PPT file and find the SmartArt shapes.
st = []
st.append("Below is extracted text from SmartArt:")
for i, unusedItem in enumerate(presentation.Slides):
    for j, unusedItem in enumerate(presentation.Slides[i].Shapes):
        if isinstance(presentation.Slides[i].Shapes[j], ISmartArt):
            smartArt = presentation.Slides[i].Shapes[j] if isinstance(
                presentation.Slides[i].Shapes[j], ISmartArt) else None

            # Extract text from SmartArt and append to the StringBuilder object.
            for k, unusedItem in enumerate(smartArt.Nodes):
                st.append(smartArt.Nodes[k].TextFrame.Text)

# Save to file.
AppendAllText(outputFile, st)
presentation.Dispose()
