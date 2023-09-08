from spire.presentation.common import *
from spire.presentation import *

inputFile = "./Data/ChangeHyperlinkColor.pptx"
outputFile = "ChangeHyperlinkColor.pptx"

presentation = Presentation()
presentation.LoadFromFile(inputFile)

#Get the first slide
slide = presentation.Slides[0]

#Get the theme of the slide
theme = slide.Theme

#Change the color of hyperlink to red
theme.ColorScheme.HyperlinkColor.Color = Color.get_Red()
#Save to file.
presentation.SaveToFile(outputFile, FileFormat.Pptx2013)
presentation.Dispose()

        


    
